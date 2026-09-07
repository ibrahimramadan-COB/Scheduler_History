"""
Takes the cleaned (all-event-types) DataFrame from cleaner.py, joins in
GROUP / GROUP 2 by matching CREATOR USER against the 'Names' tab's
'Scheduler Name' column, and UPSERTS into the 'Master Data' tab: for
any row whose A-I key (APPOINTMENT DATE, APPOINTMENT TYPE, PATIENT,
CASE, CLINIC, CALENDAR NAME, CREATOR USER, CREATED TIMESTAMP,
EVENT ACTION) already exists, the FRESH scrape's version always wins
— every column, not just the key fields. This is what guarantees
Master Data reflects the newest scrape, never a stale earlier one.

Auth: a Google service account JSON key, passed via the
GOOGLE_SERVICE_ACCOUNT_JSON env var (the *contents* of the key file,
not a path — GitHub Actions secrets are just text). That service
account's email must be shared as an Editor on the target Sheet.
"""

import os
import json
import time
from datetime import datetime

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

MASTER_TAB = "Master Data"
NAMES_TAB = "Names"
RUNS_TAB = "Runs"
CLINICS_TAB = "Clinics"

MASTER_DATA_COLUMNS = [
    "APPOINTMENT DATE", "APPOINTMENT TYPE", "PATIENT", "CASE", "CLINIC",
    "CALENDAR NAME", "CREATOR USER", "CREATED TIMESTAMP", "EVENT ACTION",
    "DETAILS", "GROUP", "GROUP 2", "SOURCE_FILE", "DATA_SOURCE_TAB",
]

# The TRUE identity of an event — columns A-I only. Anything scraped again
# with the same values here IS the same real-world event; whichever scrape
# produced it most recently should win on every OTHER column (DETAILS,
# GROUP, GROUP 2, SOURCE_FILE, DATA_SOURCE_TAB).
DEDUPE_KEY = [
    "APPOINTMENT DATE", "APPOINTMENT TYPE", "PATIENT", "CASE",
    "CLINIC", "CALENDAR NAME", "CREATOR USER", "CREATED TIMESTAMP", "EVENT ACTION",
]

# Retry settings for transient Google API errors (e.g. 500/503)
MAX_RETRIES = 5
RETRY_BASE_DELAY_SECONDS = 2


def _with_retry(fn, *args, description="Google API call", **kwargs):
    """
    Calls fn(*args, **kwargs), retrying with exponential backoff on
    transient gspread.exceptions.APIError (5xx). Raises immediately on
    non-transient errors (4xx) since retrying those won't help.
    """
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


def _truncate_timestamp_to_date(df: pd.DataFrame) -> pd.DataFrame:
    """
    Drops the time-of-day from CREATED TIMESTAMP, keeping only the date.
    This has to happen BEFORE dedupe/upsert — otherwise a row already
    stored with a truncated timestamp (from a prior maintenance pass)
    would never match the same event scraped again with a full
    timestamp, breaking both dedupe AND "newest wins".
    """
    if "CREATED TIMESTAMP" not in df.columns or df.empty:
        return df
    df = df.copy()
    parsed = pd.to_datetime(df["CREATED TIMESTAMP"], errors="coerce")
    date_only = parsed.dt.normalize()
    # Where parsing failed, keep the original value rather than losing data
    df["CREATED TIMESTAMP"] = date_only.where(parsed.notna(), df["CREATED TIMESTAMP"])
    return df


def get_client() -> gspread.Client:
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON env var is not set.")
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


def load_name_group_lookup(spreadsheet) -> dict:
    """Returns {normalized Scheduler Name: (Group, Group 2)} from the Names tab."""
    ws = spreadsheet.worksheet(NAMES_TAB)
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


def update_clinics_tab(spreadsheet, all_clinics: list[str]) -> None:
    """
    Overwrites the 'Clinics' tab with the FULL list of clinics WebPT reported
    on this run (regardless of whether this particular run was scoped to a
    subset via CLINIC_NAMES). This is what the web app's clinic multi-select
    reads from, so it stays accurate even after a partial/filtered run.
    """
    if not all_clinics:
        return
    ws = get_or_create_worksheet(spreadsheet, CLINICS_TAB, header=["Clinic"])
    _with_retry(ws.clear, description="clear Clinics tab")
    _with_retry(ws.append_row, ["Clinic"], value_input_option="RAW",
                description="write Clinics header")
    _with_retry(ws.append_rows, [[c] for c in sorted(all_clinics)], value_input_option="RAW",
                description="write Clinics rows")
    print(f"   🏥 Refreshed '{CLINICS_TAB}' tab with {len(all_clinics)} clinic(s).")


def write_master_data_safely(spreadsheet, live_ws, df: pd.DataFrame) -> None:
    """
    Writes the FULL final Master Data content via a safe temp-sheet-swap:
    write to a brand-new temp worksheet, verify the row count landed
    correctly, and ONLY THEN delete the live sheet and rename the temp
    sheet into its place. A transient API failure mid-write can never
    leave 'Master Data' empty or half-written — it stays exactly as it
    was until the replacement is fully verified.
    """
    temp_title = f"{MASTER_TAB}_NEW_{int(time.time())}"
    temp_ws = None
    try:
        temp_ws = _with_retry(
            spreadsheet.add_worksheet, title=temp_title,
            rows=max(len(df) + 10, 100), cols=max(len(MASTER_DATA_COLUMNS), 10),
            description="create temp Master Data worksheet",
        )
        _with_retry(temp_ws.append_row, MASTER_DATA_COLUMNS, value_input_option="RAW",
                    description="write temp header")
        if not df.empty:
            values_out = df.astype(object).where(pd.notnull(df), "").values.tolist()
            _with_retry(temp_ws.append_rows, values_out, value_input_option="USER_ENTERED",
                        description="write temp rows")

        check_values = _with_retry(temp_ws.get_all_values, description="verify temp sheet")
        actual_row_count = max(len(check_values) - 1, 0)  # minus header
        if actual_row_count != len(df):
            raise RuntimeError(
                f"Verification failed: temp sheet has {actual_row_count} data row(s), "
                f"expected {len(df)}. Aborting swap — '{MASTER_TAB}' left untouched."
            )

        _with_retry(spreadsheet.del_worksheet, live_ws, description="delete old Master Data")
        _with_retry(temp_ws.update_title, MASTER_TAB, description="rename temp sheet to Master Data")

    except Exception:
        if temp_ws is not None:
            try:
                spreadsheet.del_worksheet(temp_ws)
            except Exception:
                pass
        raise


def write_to_sheet(cleaned_df: pd.DataFrame, start_date: str, end_date: str) -> dict:
    sheet_id = os.environ.get("SHEET_ID")
    if not sheet_id:
        raise RuntimeError("SHEET_ID env var is not set.")

    client = get_client()
    spreadsheet = client.open_by_key(sheet_id)

    lookup = load_name_group_lookup(spreadsheet)
    cleaned_df = apply_group_lookup(cleaned_df, lookup)
    cleaned_df = _truncate_timestamp_to_date(cleaned_df)

    for col in MASTER_DATA_COLUMNS:
        if col not in cleaned_df.columns:
            cleaned_df[col] = None
    cleaned_df = cleaned_df[MASTER_DATA_COLUMNS]

    master_ws = get_or_create_worksheet(spreadsheet, MASTER_TAB, header=MASTER_DATA_COLUMNS)

    existing_records = _with_retry(master_ws.get_all_records, description="read existing Master Data")
    existing_df = pd.DataFrame(existing_records) if existing_records else pd.DataFrame(columns=MASTER_DATA_COLUMNS)
    existing_df = _truncate_timestamp_to_date(existing_df)
    existing_count = len(existing_df)
    print(f"   📥 Master Data currently has {existing_count} row(s) on record.")

    for col in MASTER_DATA_COLUMNS:
        if col not in existing_df.columns:
            existing_df[col] = None
    existing_df = existing_df[MASTER_DATA_COLUMNS] if not existing_df.empty else existing_df

    # Existing rows FIRST, this run's fresh scrape LAST — so keep="last"
    # always prefers whatever the most recent scrape produced for any
    # event that appears in both.
    combined = pd.concat([existing_df, cleaned_df], ignore_index=True, sort=False)
    for col in MASTER_DATA_COLUMNS:
        if col not in combined.columns:
            combined[col] = None
    combined = combined[MASTER_DATA_COLUMNS]

    before = len(combined)
    combined.drop_duplicates(subset=DEDUPE_KEY, keep="last", inplace=True)
    after = len(combined)
    removed = before - after

    # For the run log: how many of this run's rows were genuinely new
    # (as opposed to updates to something that already existed)
    new_keys = set(
        tuple(str(v) for v in row) for row in cleaned_df[DEDUPE_KEY].itertuples(index=False, name=None)
    )
    existing_keys = set(
        tuple(str(v) for v in row) for row in existing_df[DEDUPE_KEY].itertuples(index=False, name=None)
    ) if not existing_df.empty else set()
    genuinely_new = len(new_keys - existing_keys)
    updated_existing = len(new_keys & existing_keys)

    write_master_data_safely(spreadsheet, master_ws, combined)
    # get_or_create_worksheet's cached ref is now stale (old sheet deleted) —
    # re-fetch for anything below that might need it in the future.
    master_ws = spreadsheet.worksheet(MASTER_TAB)

    print(f"   ➕ {genuinely_new} new row(s), 🔄 {updated_existing} existing row(s) refreshed with newer data, "
          f"🧹 {removed} exact-duplicate row(s) collapsed. Master Data now has {after} row(s).")

    runs_ws = get_or_create_worksheet(
        spreadsheet, RUNS_TAB,
        header=["Timestamp", "Start Date", "End Date", "Rows Scraped", "New Rows",
                "Existing Rows Refreshed", "Duplicates Collapsed", "Total Rows in Master"],
    )
    _with_retry(
        runs_ws.append_row,
        [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            start_date,
            end_date,
            len(cleaned_df),
            genuinely_new,
            updated_existing,
            removed,
            after,
        ],
        value_input_option="USER_ENTERED",
        description="log run to Runs tab",
    )

    all_clinics_raw = os.environ.get("ALL_CLINICS", "")
    if all_clinics_raw:
        update_clinics_tab(spreadsheet, [c for c in all_clinics_raw.split("|") if c.strip()])

    return {
        "rows_scraped": len(cleaned_df),
        "new_rows": genuinely_new,
        "updated_rows": updated_existing,
        "duplicates_collapsed": removed,
        "total_rows": after,
    }


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
    print(f"\n✅ Done. Scraped {result['rows_scraped']} rows: {result['new_rows']} new, "
          f"{result['updated_rows']} refreshed with newer data, {result['duplicates_collapsed']} duplicate(s) collapsed. "
          f"Master Data now has {result['total_rows']} rows.")
