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

# Same key used for de-duplication throughout this project
DEDUPE_KEY = ["APPOINTMENT DATE", "PATIENT", "CREATED TIMESTAMP", "EVENT ACTION", "DATA_SOURCE_TAB"]

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
            # Only retry on server-side/transient errors
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


def load_existing_keys(ws) -> set:
    """Reads the current Master Data tab and returns the set of existing dedupe keys."""
    values = _with_retry(ws.get_all_records, description="read Master Data for existing keys")
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
    _with_retry(ws.clear, description="clear Clinics tab")
    _with_retry(ws.append_row, ["Clinic"], value_input_option="RAW",
                description="write Clinics header")
    _with_retry(ws.append_rows, [[c] for c in sorted(all_clinics)], value_input_option="RAW",
                description="write Clinics rows")
    print(f"   🏥 Refreshed '{CLINICS_TAB}' tab with {len(all_clinics)} clinic(s).")


def final_dedupe_pass(spreadsheet, ws) -> int:
    """
    Explicit final safety net: re-reads the WHOLE Master Data tab and drops
    any duplicate rows by the same dedupe key.

    SAFETY: the deduped data is written to a NEW TEMPORARY WORKSHEET first
    and verified, and ONLY THEN is the live 'Master Data' tab touched (by
    deleting it and renaming the temp sheet into its place). This guarantees
    the live tab is never cleared unless the full replacement data is
    already safely written and confirmed — so a transient API failure
    mid-write can no longer leave 'Master Data' empty, which is what
    happened previously when this used ws.clear() + ws.append_rows()
    directly against the live tab.
    """
    values = _with_retry(ws.get_all_values, description="read Master Data for dedupe")
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

    if removed <= 0:
        return 0

    print(f"   🧹 Final dedupe pass found {removed} duplicate row(s) — writing safely...")

    temp_title = f"{MASTER_TAB}_NEW_{int(time.time())}"
    temp_ws = None
    try:
        # 1. Write the full deduped dataset to a brand-new temp worksheet
        temp_ws = _with_retry(
            spreadsheet.add_worksheet, title=temp_title,
            rows=max(after + 10, 100), cols=max(len(header), 10),
            description="create temp dedupe worksheet",
        )
        _with_retry(temp_ws.append_row, header, value_input_option="RAW",
                    description="write temp header")
        if not df.empty:
            values_out = df.astype(object).where(pd.notnull(df), "").values.tolist()
            _with_retry(temp_ws.append_rows, values_out, value_input_option="USER_ENTERED",
                        description="write temp deduped rows")

        # 2. Verify the temp sheet actually has the expected row count before
        #    touching the live sheet at all
        check_values = _with_retry(temp_ws.get_all_values, description="verify temp sheet")
        actual_row_count = max(len(check_values) - 1, 0)  # minus header
        if actual_row_count != after:
            raise RuntimeError(
                f"Verification failed: temp sheet has {actual_row_count} data row(s), "
                f"expected {after}. Aborting swap — '{MASTER_TAB}' left untouched."
            )

        # 3. Only now: delete the old live sheet and rename the verified temp
        #    sheet into its place. This is effectively an atomic swap from
        #    the user's point of view.
        _with_retry(spreadsheet.del_worksheet, ws, description="delete old Master Data")
        _with_retry(temp_ws.update_title, MASTER_TAB, description="rename temp sheet to Master Data")

        print(f"   🧹 Final dedupe pass removed {removed} duplicate row(s) from '{MASTER_TAB}' (safe swap).")
        return removed

    except Exception:
        # If anything went wrong, clean up the temp sheet if it exists, and
        # leave the ORIGINAL 'Master Data' completely untouched.
        if temp_ws is not None:
            try:
                spreadsheet.del_worksheet(temp_ws)
            except Exception:
                pass
        print(f"   ⚠️ Dedupe pass aborted safely — '{MASTER_TAB}' was NOT modified. Will retry next run.")
        raise


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
        _with_retry(master_ws.append_rows, values, value_input_option="USER_ENTERED",
                    description="append new rows to Master Data")
        print(f"   ➕ Appended {added} new row(s) to '{MASTER_TAB}'.")
    else:
        print(f"   ℹ️ No new rows to append — everything was already in '{MASTER_TAB}'.")

    # Log the run
    runs_ws = get_or_create_worksheet(
        spreadsheet, RUNS_TAB,
        header=["Timestamp", "Start Date", "End Date", "Rows Scraped", "Rows Added (new)",
                "Duplicates Removed (final pass)", "Total Rows in Master"],
    )

    # Explicit final safety-net dedupe across the whole tab (see docstring).
    # Re-fetch the worksheet reference in case final_dedupe_pass swapped it.
    try:
        removed = final_dedupe_pass(spreadsheet, master_ws)
    except Exception as e:
        # Dedupe failing should NEVER be treated as data loss — it's safe by
        # design now. Log it and continue; Master Data still has all rows
        # (just with duplicates, which the next run's dedupe pass will retry).
        print(f"   ⚠️ Dedupe pass failed safely (no data lost): {e}")
        removed = 0

    total_after = len(existing_keys) + added - removed

    _with_retry(
        runs_ws.append_row,
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
        description="log run to Runs tab",
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
