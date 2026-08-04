"""
Takes the cleaned (all-event-types) DataFrame from cleaner.py, joins in
GROUP / GROUP 2 by matching CREATOR USER against the 'Names' tab's
'Scheduler Name' column, appends any genuinely new rows to the 'Master
Data' tab in the Google Sheet (deduped against what's already there),
and appends one row to the 'Runs' tab logging this run's date range.

Auth: a Google service account JSON key, passed via the
GOOGLE_SERVICE_ACCOUNT_JSON env var (the *contents* of the key file,
not a path — GitHub Actions secrets are just text). That service
account's email must be shared as an Editor on the target Sheet.
"""

import os
import json
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

# Same key used for de-duplication throughout this project
DEDUPE_KEY = ["APPOINTMENT DATE", "PATIENT", "CREATED TIMESTAMP", "EVENT ACTION", "DATA_SOURCE_TAB"]


def _normalize_name(name) -> str:
    if name is None:
        return ""
    return " ".join(str(name).strip().lower().split())


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
        ws = spreadsheet.add_worksheet(title=title, rows=1000, cols=max(len(header or []), 10))
        if header:
            ws.append_row(header, value_input_option="RAW")
    return ws


def load_name_group_lookup(spreadsheet) -> dict:
    """Returns {normalized Scheduler Name: (Group, Group 2)} from the Names tab."""
    ws = spreadsheet.worksheet(NAMES_TAB)
    records = ws.get_all_records()  # uses row 1 as header
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


def load_existing_keys(ws) -> set:
    """Reads the current Master Data tab and returns the set of existing dedupe keys."""
    values = ws.get_all_records()
    if not values:
        return set()
    existing_df = pd.DataFrame(values)
    missing_cols = [c for c in DEDUPE_KEY if c not in existing_df.columns]
    if missing_cols:
        return set()
    keys = set(
        tuple(str(v) for v in row)
        for row in existing_df[DEDUPE_KEY].itertuples(index=False, name=None)
    )
    return keys


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
    ws.clear()
    ws.append_row(["Clinic"], value_input_option="RAW")
    ws.append_rows([[c] for c in sorted(all_clinics)], value_input_option="RAW")
    print(f"   🏥 Refreshed '{CLINICS_TAB}' tab with {len(all_clinics)} clinic(s).")


def final_dedupe_pass(ws) -> int:
    """
    Explicit final safety net: re-reads the WHOLE Master Data tab and drops
    any duplicate rows by the same dedupe key, rewriting the tab if any were
    found. The append step already skips rows that match something already
    present, so this should normally be a no-op — it exists to guarantee
    correctness even if something ever got in another way (a manual paste,
    a race between two runs, etc.), per explicit request: "at the end we
    drop duplicates to make sure everything is accurate."
    """
    values = ws.get_all_values()
    if len(values) < 2:
        return 0
    header, data_rows = values[0], values[1:]
    df = pd.DataFrame(data_rows, columns=header)

    missing_cols = [c for c in DEDUPE_KEY if c not in df.columns]
    if missing_cols:
        return 0

    before = len(df)
    df.drop_duplicates(subset=DEDUPE_KEY, keep="first", inplace=True)
    after = len(df)
    removed = before - after
    if removed > 0:
        print(f"   🧹 Final dedupe pass removed {removed} duplicate row(s) from '{MASTER_TAB}'.")
        ws.clear()
        ws.append_row(header, value_input_option="RAW")
        if not df.empty:
            values_out = df.astype(object).where(pd.notnull(df), "").values.tolist()
            ws.append_rows(values_out, value_input_option="USER_ENTERED")
    return removed


def write_to_sheet(cleaned_df: pd.DataFrame, start_date: str, end_date: str) -> dict:
    sheet_id = os.environ.get("SHEET_ID")
    if not sheet_id:
        raise RuntimeError("SHEET_ID env var is not set.")

    client = get_client()
    spreadsheet = client.open_by_key(sheet_id)

    lookup = load_name_group_lookup(spreadsheet)
    cleaned_df = apply_group_lookup(cleaned_df, lookup)

    # Make sure every expected column exists, in the right order
    for col in MASTER_DATA_COLUMNS:
        if col not in cleaned_df.columns:
            cleaned_df[col] = None
    cleaned_df = cleaned_df[MASTER_DATA_COLUMNS]

    master_ws = get_or_create_worksheet(spreadsheet, MASTER_TAB, header=MASTER_DATA_COLUMNS)
    existing_keys = load_existing_keys(master_ws)
    print(f"   📥 Master Data currently has {len(existing_keys)} row(s) on record.")

    def row_key(row):
        return tuple(str(row[c]) for c in DEDUPE_KEY)

    if not cleaned_df.empty:
        is_new = cleaned_df.apply(lambda r: row_key(r) not in existing_keys, axis=1)
        new_rows_df = cleaned_df[is_new]
    else:
        new_rows_df = cleaned_df

    added = len(new_rows_df)
    if added:
        # Sheets can't hold NaN/NaT — normalize to strings/blank first
        values = new_rows_df.astype(object).where(pd.notnull(new_rows_df), "").values.tolist()
        master_ws.append_rows(values, value_input_option="USER_ENTERED")
        print(f"   ➕ Appended {added} new row(s) to '{MASTER_TAB}'.")
    else:
        print(f"   ℹ️ No new rows to append — everything was already in '{MASTER_TAB}'.")

    # Log the run
    runs_ws = get_or_create_worksheet(
        spreadsheet, RUNS_TAB,
        header=["Timestamp", "Start Date", "End Date", "Rows Scraped", "Rows Added (new)",
                "Duplicates Removed (final pass)", "Total Rows in Master"],
    )
    # Explicit final safety-net dedupe across the whole tab (see docstring)
    removed = final_dedupe_pass(master_ws)
    total_after = len(existing_keys) + added - removed

    runs_ws.append_row(
        [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            start_date,
            end_date,
            len(cleaned_df),
            added,
            removed,
            total_after,
        ],
        value_input_option="USER_ENTERED",
    )

    all_clinics_raw = os.environ.get("ALL_CLINICS", "")
    if all_clinics_raw:
        update_clinics_tab(spreadsheet, [c for c in all_clinics_raw.split("|") if c.strip()])

    return {"rows_scraped": len(cleaned_df), "rows_added": added, "duplicates_removed": removed, "total_rows": total_after}


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
    print(f"\n✅ Done. Scraped {result['rows_scraped']} rows, added {result['rows_added']} new, "
          f"removed {result['duplicates_removed']} duplicate(s) in final pass, "
          f"Master Data now has {result['total_rows']} rows.")
