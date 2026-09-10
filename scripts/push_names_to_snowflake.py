"""
Pushes the Names tab (Scheduler Name, Group, Group 2) from the main
Google Sheet into Snowflake's SCHEDULER_HISTORY_NAMES table — a full
TRUNCATE + INSERT every time, since this table is small and is always
meant to exactly mirror the current sheet, not accumulate incrementally.

Triggered manually via the Apps Script "Push Names to Snowflake" menu
item, which calls this repo's push-names.yml workflow. Not part of the
scrape pipeline — runs independently, whenever you actually edit Names.
"""

import os
import sys
import json

import gspread
from google.oauth2.service_account import Credentials

from snowflake_writer import get_connection  # reuse the same auth/connection logic

SHEET_SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
NAMES_TAB = "Names"
TABLE_NAME = "SCHEDULER_HISTORY_NAMES"


def get_sheets_client():
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON env var is not set.")
    creds = Credentials.from_service_account_info(json.loads(raw), scopes=SHEET_SCOPES)
    return gspread.authorize(creds)


def read_names_from_sheet(sheet_id: str) -> list[tuple]:
    client = get_sheets_client()
    spreadsheet = client.open_by_key(sheet_id)
    ws = spreadsheet.worksheet(NAMES_TAB)

    # Read raw values instead of get_all_records() — the header row has a
    # duplicate 'Group'/'Group 2' column somewhere, which get_all_records()
    # refuses to handle. Taking the FIRST occurrence of each needed header
    # name sidesteps that without requiring the sheet itself to be cleaned up.
    all_values = ws.get_all_values()
    if not all_values:
        return []
    header = all_values[0]
    data_rows = all_values[1:]

    def first_index(name):
        for i, h in enumerate(header):
            if h.strip() == name:
                return i
        return None

    idx_name = first_index("Scheduler Name")
    idx_group = first_index("Group")
    idx_group2 = first_index("Group 2")
    if idx_name is None:
        raise RuntimeError(f"'{NAMES_TAB}' tab has no 'Scheduler Name' column. Headers found: {header}")

    rows = []
    for r in data_rows:
        scheduler_name = str(r[idx_name]).strip() if idx_name < len(r) else ""
        if not scheduler_name:
            continue
        group_name = str(r[idx_group]).strip() if idx_group is not None and idx_group < len(r) else ""
        group_2 = str(r[idx_group2]).strip() if idx_group2 is not None and idx_group2 < len(r) else ""
        rows.append((scheduler_name, group_name, group_2))
    return rows


def push_to_snowflake(rows: list[tuple]) -> int:
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(f"TRUNCATE TABLE {TABLE_NAME}")
        if rows:
            cursor.executemany(
                f"INSERT INTO {TABLE_NAME} (SCHEDULER_NAME, GROUP_NAME, GROUP_2) VALUES (%s, %s, %s)",
                rows,
            )
        return len(rows)
    finally:
        cursor.close()
        conn.close()


if __name__ == "__main__":
    sheet_id = os.environ.get("SHEET_ID")
    if not sheet_id:
        print("❌ SHEET_ID env var is not set.")
        sys.exit(1)

    print(f"📥 Reading '{NAMES_TAB}' tab...")
    rows = read_names_from_sheet(sheet_id)
    print(f"   Found {len(rows)} row(s) with a Scheduler Name.")

    print(f"📤 Pushing to {TABLE_NAME} (full replace)...")
    count = push_to_snowflake(rows)
    print(f"✅ Done. {TABLE_NAME} now has {count} row(s).")
