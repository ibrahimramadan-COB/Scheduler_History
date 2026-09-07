"""
Routes data to the correct per-month spreadsheet file, based on
APPOINTMENT DATE's month. Reads/maintains the 'Month Files' index tab
in the main control sheet, and can auto-create a new monthly file the
first time a month is needed (going past what the initial migration
registered).
"""

import os
import datetime

import gspread

MONTH_FILES_TAB = "Month Files"
MONTHLY_MASTER_TAB = "Master Data"


def ensure_month_files_tab(main_spreadsheet):
    try:
        ws = main_spreadsheet.worksheet(MONTH_FILES_TAB)
    except gspread.exceptions.WorksheetNotFound:
        ws = main_spreadsheet.add_worksheet(title=MONTH_FILES_TAB, rows=200, cols=3)
        ws.append_row(["Month Key", "Display Name", "Spreadsheet ID"], value_input_option="RAW")
    # Force column A to plain text — otherwise Sheets can silently auto-convert
    # a value like "2026-08" into an actual Date, which breaks every lookup
    # against it (this exact bug bit the Apps Script side of this project).
    ws.format("A:A", {"numberFormat": {"type": "TEXT"}})
    return ws


def read_month_files_index(main_spreadsheet) -> dict:
    """Returns {month_key: {"display": ..., "spreadsheet_id": ...}}."""
    ws = ensure_month_files_tab(main_spreadsheet)
    values = ws.get_all_values()
    index = {}
    for row in values[1:]:
        if not row or not row[0]:
            continue
        key = str(row[0]).strip()
        index[key] = {
            "display": row[1].strip() if len(row) > 1 and row[1] else key,
            "spreadsheet_id": row[2].strip() if len(row) > 2 else None,
        }
    return index


def display_name_for_month_key(month_key: str) -> str:
    y, m = month_key.split("-")
    d = datetime.date(int(y), int(m), 1)
    return d.strftime("%B %Y")


def get_or_create_month_file(client, main_spreadsheet, month_key: str,
                              share_email: str | None, folder_id: str | None) -> dict:
    """
    Returns {"display": ..., "spreadsheet_id": ...} for the given month,
    creating a brand-new spreadsheet file (and registering it in Month
    Files) if one doesn't exist yet.
    """
    index = read_month_files_index(main_spreadsheet)
    if month_key in index and index[month_key]["spreadsheet_id"]:
        return index[month_key]

    display = display_name_for_month_key(month_key)
    print(f"   🆕 No monthly file registered for {month_key} — creating '{display}'...")

    create_kwargs = {}
    if folder_id:
        create_kwargs["folder_id"] = folder_id
    try:
        new_ss = client.create(f"WebPT Master Data — {display}", **create_kwargs)
    except gspread.exceptions.APIError as e:
        raise RuntimeError(
            f"Could not create a new spreadsheet for {display}. This is often a Drive storage-quota "
            f"issue for service accounts — set MONTHLY_FILES_FOLDER_ID to a Shared Drive folder the "
            f"service account has access to. Original error: {e}"
        )

    if share_email:
        try:
            new_ss.share(share_email, perm_type="user", role="writer")
        except Exception as e:
            print(f"   ⚠️ Could not share new file with {share_email}: {e}")
    else:
        print(f"   ⚠️ MONTHLY_FILE_SHARE_EMAIL not set — this new file is owned by the service account "
              f"and won't show up in anyone's Drive until manually shared. URL: {new_ss.url}")

    ws = new_ss.sheet1
    ws.update_title(MONTHLY_MASTER_TAB)

    ws_idx = ensure_month_files_tab(main_spreadsheet)
    ws_idx.append_row([month_key, display, new_ss.id], value_input_option="RAW")

    return {"display": display, "spreadsheet_id": new_ss.id}
