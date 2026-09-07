"""
Takes the cleaned (all-event-types) DataFrame from cleaner.py, joins in
GROUP / GROUP 2 from the main sheet's Names tab, splits rows by
APPOINTMENT DATE's month, and upserts each group into its own monthly
spreadsheet file (auto-created via month_router if it doesn't exist
yet). For every monthly file touched, also updates the 'Created
Coverage Index' tab in the main sheet — a Created-Month x
Appointment-Month matrix that lets the web app know which monthly
files actually contain data for a given Created Date Range, without
having to open every file every time.

Auth: a Google service account JSON key, passed via the
GOOGLE_SERVICE_ACCOUNT_JSON env var (contents, not a path) or
GOOGLE_SERVICE_ACCOUNT_FILE (a path, for local runs). That service
account must be shared as Editor on the main control sheet, and on
every monthly file (existing ones now, and any it creates itself
going forward).
"""

import os
import json
import time
from datetime import datetime

import pandas as pd
import gspread
from gspread.utils import rowcol_to_a1
from google.oauth2.service_account import Credentials

from month_router import (
    read_month_files_index, get_or_create_month_file,
    display_name_for_month_key, MONTHLY_MASTER_TAB,
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",  # needed to auto-create new monthly files
]

NAMES_TAB = "Names"
RUNS_TAB = "Runs"
CLINICS_TAB = "Clinics"
COVERAGE_INDEX_TAB = "Created Coverage Index"

MASTER_DATA_COLUMNS = [
    "APPOINTMENT DATE", "APPOINTMENT TYPE", "PATIENT", "CASE", "CLINIC",
    "CALENDAR NAME", "CREATOR USER", "CREATED TIMESTAMP", "EVENT ACTION",
    "DETAILS", "GROUP", "GROUP 2", "SOURCE_FILE", "DATA_SOURCE_TAB",
]

DEDUPE_KEY = [
    "APPOINTMENT DATE", "APPOINTMENT TYPE", "PATIENT", "CASE",
    "CLINIC", "CALENDAR NAME", "CREATOR USER", "CREATED TIMESTAMP", "EVENT ACTION",
]

MAX_RETRIES = 5
RETRY_BASE_DELAY_SECONDS = 2
WRITE_CHUNK_SIZE = 5000


def _with_retry(fn, *args, description="Google API call", **kwargs):
    last_exc = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            last_exc = e
            status = None
            try:
                status = e.response.status_code
            except Exception:
                pass
            if status is not None and status < 500:
                raise
            if attempt == MAX_RETRIES:
                break
            delay = RETRY_BASE_DELAY_SECONDS * (2 ** (attempt - 1))
            print(f"   ⚠️ {description} failed (attempt {attempt}/{MAX_RETRIES}, status={status}). "
                  f"Retrying in {delay}s...")
            time.sleep(delay)
    raise last_exc


def _normalize_name(name) -> str:
    if name is None:
        return ""
    return " ".join(str(name).strip().lower().split())


def get_client() -> gspread.Client:
    file_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE")
    if file_path:
        if not os.path.exists(file_path):
            raise RuntimeError(f"GOOGLE_SERVICE_ACCOUNT_FILE points to a file that doesn't exist: {file_path}")
        creds = Credentials.from_service_account_file(file_path, scopes=SCOPES)
        return gspread.authorize(creds)

    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        raise RuntimeError("Set either GOOGLE_SERVICE_ACCOUNT_FILE (path) or GOOGLE_SERVICE_ACCOUNT_JSON (content).")
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def get_or_create_worksheet(spreadsheet, title: str, header: list[str] | None = None):
    try:
        ws = spreadsheet.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = _with_retry(
            spreadsheet.add_worksheet, title=title, rows=1000, cols=max(len(header or []), 10),
            description=f"create worksheet '{title}'",
        )
        if header:
            _with_retry(ws.append_row, header, value_input_option="RAW",
                        description=f"write header to '{title}'")
    return ws


def load_name_group_lookup(main_spreadsheet) -> dict:
    ws = main_spreadsheet.worksheet(NAMES_TAB)
    records = _with_retry(ws.get_all_records, description="read Names tab")
    lookup = {}
    for r in records:
        scheduler_name = r.get("Scheduler Name")
        if not scheduler_name:
            continue
        key = _normalize_name(scheduler_name)
        lookup[key] = (r.get("Group") or "Other", r.get("Group 2") or "Other")
    print(f"   📇 Loaded {len(lookup)} scheduler-name lookups from '{NAMES_TAB}' tab.")
    return lookup


def apply_group_lookup(df: pd.DataFrame, lookup: dict) -> pd.DataFrame:
    if df.empty:
        df["GROUP"] = []
        df["GROUP 2"] = []
        return df
    normalized = df["CREATOR USER"].apply(_normalize_name)
    groups = normalized.map(lambda n: lookup.get(n, ("Other", "Other")))
    df = df.copy()
    df["GROUP"] = groups.apply(lambda g: g[0])
    df["GROUP 2"] = groups.apply(lambda g: g[1])
    unmatched = int((df["GROUP"] == "Other").sum())
    if unmatched:
        print(f"   ℹ️ {unmatched} row(s) had a CREATOR USER not found in '{NAMES_TAB}' — tagged 'Other'.")
    return df


def _truncate_timestamp_to_date(df: pd.DataFrame) -> pd.DataFrame:
    if "CREATED TIMESTAMP" not in df.columns or df.empty:
        return df
    df = df.copy()
    parsed = pd.to_datetime(df["CREATED TIMESTAMP"], errors="coerce")
    date_only = parsed.dt.normalize()
    df["CREATED TIMESTAMP"] = date_only.where(parsed.notna(), df["CREATED TIMESTAMP"])
    return df


def month_key_series(date_series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(date_series, errors="coerce")
    return parsed.dt.strftime("%Y-%m")


def _canonical_date_str(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce")
    return parsed.dt.strftime("%Y-%m-%d").where(parsed.notna(), series.astype(str).str.strip())


def build_dedupe_keys(df: pd.DataFrame) -> pd.Series:
    parts = []
    for col in DEDUPE_KEY:
        if col in ("APPOINTMENT DATE", "CREATED TIMESTAMP"):
            parts.append(_canonical_date_str(df[col]))
        else:
            parts.append(df[col].astype(str).str.strip())
    key_df = pd.concat(parts, axis=1)
    return key_df.astype(str).agg("‖".join, axis=1)


def update_clinics_tab(main_spreadsheet, all_clinics: list[str]) -> None:
    if not all_clinics:
        return
    ws = get_or_create_worksheet(main_spreadsheet, CLINICS_TAB, header=["Clinic"])
    _with_retry(ws.clear, description="clear Clinics tab")
    _with_retry(ws.append_row, ["Clinic"], value_input_option="RAW", description="write Clinics header")
    _with_retry(ws.append_rows, [[c] for c in sorted(all_clinics)], value_input_option="RAW",
                description="write Clinics rows")
    print(f"   🏥 Refreshed '{CLINICS_TAB}' tab with {len(all_clinics)} clinic(s).")


def write_sheet_in_place(live_ws, header: list[str], df: pd.DataFrame) -> None:
    num_cols = len(header)
    values_out = df.astype(object).where(pd.notnull(df), "").values.tolist() if not df.empty else []
    all_rows = [header] + values_out
    new_row_count = len(all_rows)

    old_row_count = live_ws.row_count
    old_col_count = live_ws.col_count

    target_row_count = max(old_row_count, new_row_count)
    target_col_count = max(old_col_count, num_cols)
    if target_row_count != old_row_count or target_col_count != old_col_count:
        _with_retry(live_ws.resize, rows=target_row_count, cols=target_col_count,
                    description="grow sheet to fit")

    for start in range(0, new_row_count, WRITE_CHUNK_SIZE):
        chunk = all_rows[start:start + WRITE_CHUNK_SIZE]
        first_row, last_row = start + 1, start + len(chunk)
        range_name = f"A{first_row}:{rowcol_to_a1(last_row, num_cols)}"
        _with_retry(live_ws.update, values=chunk, range_name=range_name,
                    value_input_option="USER_ENTERED",
                    description=f"write rows {first_row}-{last_row}")

    if old_row_count > new_row_count:
        for start in range(new_row_count, old_row_count, WRITE_CHUNK_SIZE):
            end = min(start + WRITE_CHUNK_SIZE, old_row_count)
            blank_rows = [[""] * num_cols for _ in range(end - start)]
            range_name = f"A{start + 1}:{rowcol_to_a1(end, num_cols)}"
            _with_retry(live_ws.update, values=blank_rows, range_name=range_name,
                        value_input_option="USER_ENTERED",
                        description=f"blank leftover rows {start + 1}-{end}")

    if new_row_count < target_row_count:
        try:
            _with_retry(live_ws.resize, rows=new_row_count, cols=num_cols,
                        description="trim sheet back to size")
        except Exception as e:
            print(f"   ⚠️ Could not trim sheet dimensions after write (non-fatal): {e}")


def upsert_month_file(client, spreadsheet_id: str, new_rows_df: pd.DataFrame) -> pd.DataFrame:
    ss = client.open_by_key(spreadsheet_id)
    ws = get_or_create_worksheet(ss, MONTHLY_MASTER_TAB, header=MASTER_DATA_COLUMNS)

    existing_records = _with_retry(ws.get_all_records, description="read existing monthly file")
    existing_df = pd.DataFrame(existing_records) if existing_records else pd.DataFrame(columns=MASTER_DATA_COLUMNS)
    for col in MASTER_DATA_COLUMNS:
        if col not in existing_df.columns:
            existing_df[col] = None
    existing_df = existing_df[MASTER_DATA_COLUMNS] if not existing_df.empty else existing_df
    existing_df = _truncate_timestamp_to_date(existing_df)

    combined = pd.concat([existing_df, new_rows_df], ignore_index=True, sort=False)
    for col in MASTER_DATA_COLUMNS:
        if col not in combined.columns:
            combined[col] = None
    combined = combined[MASTER_DATA_COLUMNS]

    keys = build_dedupe_keys(combined)
    combined = combined.assign(_key=keys)
    combined = combined.sort_values("APPOINTMENT DATE", kind="stable")
    combined = combined.drop_duplicates(subset="_key", keep="last").drop(columns="_key")

    write_sheet_in_place(ws, MASTER_DATA_COLUMNS, combined)
    return combined


def compute_created_month_counts(df: pd.DataFrame) -> dict:
    if df.empty or "CREATED TIMESTAMP" not in df.columns:
        return {}
    keys = month_key_series(df["CREATED TIMESTAMP"]).dropna()
    return keys.value_counts().to_dict()


def update_coverage_index(main_spreadsheet, appointment_display: str, created_counts: dict) -> None:
    ws = get_or_create_worksheet(main_spreadsheet, COVERAGE_INDEX_TAB,
                                  header=["Created Month", "Total Number Created"])
    _with_retry(ws.format, "A:A", {"numberFormat": {"type": "TEXT"}},
                description="force Coverage Index col A to text")

    values = _with_retry(ws.get_all_values, description="read Coverage Index")
    header = list(values[0]) if values else ["Created Month", "Total Number Created"]
    data_rows = [list(r) for r in values[1:]] if len(values) > 1 else []

    column_name = f"{appointment_display} Appts"
    if column_name not in header:
        header.append(column_name)
        for r in data_rows:
            while len(r) < len(header):
                r.append("")
    col_idx = header.index(column_name)

    row_index_by_label = {r[0]: i for i, r in enumerate(data_rows)}
    for created_key in created_counts:
        label = display_name_for_month_key(created_key)
        if label not in row_index_by_label:
            new_row = [""] * len(header)
            new_row[0] = label
            data_rows.append(new_row)
            row_index_by_label[label] = len(data_rows) - 1

    label_to_count = {display_name_for_month_key(k): v for k, v in created_counts.items()}
    for r in data_rows:
        r[col_idx] = label_to_count.get(r[0], "") or ""

    for r in data_rows:
        total = 0
        for c in range(2, len(header)):
            try:
                total += int(r[c]) if r[c] not in ("", None) else 0
            except (ValueError, TypeError):
                pass
        r[1] = total

    def sort_key(r):
        try:
            return pd.to_datetime(r[0])
        except Exception:
            return pd.Timestamp.max

    data_rows.sort(key=sort_key)

    _with_retry(ws.clear, description="clear Coverage Index")
    out_rows = [header] + data_rows
    _with_retry(ws.append_rows, out_rows, value_input_option="USER_ENTERED", description="write Coverage Index")


def write_to_sheet(cleaned_df: pd.DataFrame, start_date: str, end_date: str) -> dict:
    main_sheet_id = os.environ.get("SHEET_ID")
    if not main_sheet_id:
        raise RuntimeError("SHEET_ID env var is not set.")
    share_email = os.environ.get("MONTHLY_FILE_SHARE_EMAIL")
    folder_id = os.environ.get("MONTHLY_FILES_FOLDER_ID")

    client = get_client()
    main_ss = client.open_by_key(main_sheet_id)

    lookup = load_name_group_lookup(main_ss)
    cleaned_df = apply_group_lookup(cleaned_df, lookup)
    cleaned_df = _truncate_timestamp_to_date(cleaned_df)

    for col in MASTER_DATA_COLUMNS:
        if col not in cleaned_df.columns:
            cleaned_df[col] = None
    cleaned_df = cleaned_df[MASTER_DATA_COLUMNS]

    cleaned_df = cleaned_df.assign(_month_key=month_key_series(cleaned_df["APPOINTMENT DATE"]))
    unparsed = int(cleaned_df["_month_key"].isna().sum())
    if unparsed:
        print(f"   ⚠️ {unparsed} row(s) had an unparseable APPOINTMENT DATE — skipped entirely.")
    cleaned_df = cleaned_df.dropna(subset=["_month_key"])

    per_month_summary = []

    for month_key, group in cleaned_df.groupby("_month_key"):
        group = group.drop(columns="_month_key")
        info = get_or_create_month_file(client, main_ss, month_key, share_email, folder_id)
        print(f"   📤 Writing {len(group)} scraped row(s) to {info['display']}...")

        final_df = upsert_month_file(client, info["spreadsheet_id"], group)
        print(f"   ✅ {info['display']}: {len(final_df)} row(s) now in that file (after dedupe).")

        created_counts = compute_created_month_counts(final_df)
        update_coverage_index(main_ss, info["display"], created_counts)

        per_month_summary.append(f"{info['display']}: {len(group)} scraped -> {len(final_df)} total")

    runs_ws = get_or_create_worksheet(
        main_ss, RUNS_TAB,
        header=["Timestamp", "Start Date", "End Date", "Rows Scraped", "Months Touched"],
    )
    _with_retry(
        runs_ws.append_row,
        [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            start_date, end_date, len(cleaned_df), "; ".join(per_month_summary),
        ],
        value_input_option="USER_ENTERED",
        description="log run to Runs tab",
    )

    all_clinics_raw = os.environ.get("ALL_CLINICS", "")
    if all_clinics_raw:
        update_clinics_tab(main_ss, [c for c in all_clinics_raw.split("|") if c.strip()])

    return {"rows_scraped": len(cleaned_df), "months": per_month_summary}


if __name__ == "__main__":
    import sys
    from cleaner import clean_raw_workbook

    raw_file = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("RAW_FILE")
    start_date = os.environ.get("START_DATE", "")
    end_date = os.environ.get("END_DATE", "")
    if not raw_file:
        print("❌ No raw file path given (arg or RAW_FILE env var).")
        sys.exit(1)

    df = clean_raw_workbook(raw_file)
    result = write_to_sheet(df, start_date, end_date)
    print(f"\n✅ Done. Scraped {result['rows_scraped']} rows across {len(result['months'])} month file(s):")
    for line in result["months"]:
        print(f"   - {line}")
