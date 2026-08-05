"""
Turns the raw per-clinic workbook produced by scraper.py into a flat
table of EVERY event-log row (Created, Updated, Cancelled, Checked In,
Checked Out, Deleted, Edited, No Show — whatever WebPT logged), one
row per event. Event-TYPE and department filtering happen later, in
Apps Script against the Master Data tab — but junk/non-appointment
rows (calendar blocks, placeholder patients) are stripped HERE, at the
source, so they never make it into Master Data in the first place.
"""

import os
import re
import pandas as pd

# APPOINTMENT TYPE values that aren't real appointments — internal
# calendar-block markers WebPT logs the same way as real appointments.
JUNK_APPOINTMENT_TYPES = {"blocked schedule", "calendar start", "calendar end"}

_NUMERIC_ONLY_RE = re.compile(r"^\d+$")


def is_junk_appointment_type(val) -> bool:
    if val is None:
        return False
    return str(val).strip().lower() in JUNK_APPOINTMENT_TYPES


def is_valid_patient(val) -> bool:
    """False if blank/NaN, or if the value is only digits (a placeholder ID, not a name)."""
    if val is None:
        return False
    s = str(val).strip()
    if s == "" or s.lower() == "nan":
        return False
    if _NUMERIC_ONLY_RE.match(s):
        return False
    return True


def parse_workbook(input_filepath: str) -> list[dict]:
    """Parses one raw Excel workbook (one sheet per clinic) into a list of row dicts."""
    print(f"\n   📂 Processing: {os.path.basename(input_filepath)}")
    excel_file = pd.ExcelFile(input_filepath)
    file_rows = []

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
        current_parent = None
        idx = 0

        while idx < len(df):
            row = df.iloc[idx].tolist()

            if len(row) < 3:
                idx += 1
                continue
            if pd.isna(row[0]) and pd.isna(row[1]) and pd.isna(row[2]):
                idx += 1
                continue

            row_str_0 = str(row[0]).strip()
            row_str_1 = str(row[1]).strip()
            row_str_2 = str(row[2]).strip()

            if "APPOINTMENT DATE" in row_str_0 or "Note" == row_str_0:
                idx += 1
                continue

            # A parent "appointment" row starts a new block
            if "/" in row_str_0 and len(row_str_0) <= 10 and "User" not in row_str_0:
                current_parent = {
                    "APPOINTMENT DATE": row[0],
                    "APPOINTMENT TYPE": row[1],
                    "PATIENT":          row[2],
                    "CASE":             row[3] if len(row) > 3 else None,
                    "CLINIC":           row[4] if len(row) > 4 else None,
                    "CALENDAR NAME":    row[5] if len(row) > 5 else None,
                }
                idx += 1
                continue

            if "User" in row_str_0 or "Updated Date/Time" in row_str_1:
                idx += 1
                continue

            # Every event-log row under the current parent, regardless of EVENT ACTION
            if current_parent is not None:
                file_rows.append({
                    "SOURCE_FILE":       os.path.basename(input_filepath),
                    "DATA_SOURCE_TAB":   sheet_name,
                    "APPOINTMENT DATE":  current_parent["APPOINTMENT DATE"],
                    "APPOINTMENT TYPE":  current_parent["APPOINTMENT TYPE"],
                    "PATIENT":           current_parent["PATIENT"],
                    "CASE":              current_parent["CASE"],
                    "CLINIC":            current_parent["CLINIC"],
                    "CALENDAR NAME":     current_parent["CALENDAR NAME"],
                    "CREATOR USER":      row[0],
                    "CREATED TIMESTAMP": row[1],
                    "EVENT ACTION":      row[2],
                    "DETAILS":           row[3] if len(row) > 3 else None,
                })

            idx += 1

    print(f"   ✅ Extracted {len(file_rows)} event-log rows from {len(excel_file.sheet_names)} clinic tab(s).")
    return file_rows


def clean_raw_workbook(input_filepath: str) -> pd.DataFrame:
    """Public entry point: raw workbook path -> flat, filtered, de-duplicated DataFrame (all event types)."""
    rows = parse_workbook(input_filepath)
    df = pd.DataFrame(rows)
    if df.empty:
        return df

    before_junk = len(df)
    df = df[~df["APPOINTMENT TYPE"].apply(is_junk_appointment_type)]
    df = df[df["PATIENT"].apply(is_valid_patient)]
    after_junk = len(df)
    if before_junk != after_junk:
        print(f"   🧹 Removed {before_junk - after_junk} row(s): calendar-block appointment types "
              f"(Blocked Schedule / Calendar Start / Calendar End) or non-name PATIENT values.")

    before = len(df)
    df.drop_duplicates(
        subset=["APPOINTMENT DATE", "PATIENT", "CREATED TIMESTAMP", "EVENT ACTION", "DATA_SOURCE_TAB"],
        keep="first",
        inplace=True,
    )
    after = len(df)
    if before != after:
        print(f"   🧹 Removed {before - after} duplicate rows within this file.")
    return df
