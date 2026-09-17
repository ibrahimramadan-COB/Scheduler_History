"""
scripts/cleaner.py — parses raw WebPT Scheduler History workbook into
event-log rows. Now also captures CHANGED FROM / CHANGED TO (needed to
detect recurring series and their end dates). True exact-duplicate
dedup only (not the old flawed 9-column key).

Usage: python cleaner.py <raw_input.xlsx> <cleaned_output.csv>
"""

import os
import re
import sys
import pandas as pd

JUNK_APPOINTMENT_TYPES = {"blocked schedule", "calendar start", "calendar end"}
_NUMERIC_ONLY_RE = re.compile(r"^\d+$")


def is_junk_appointment_type(val) -> bool:
    return val is not None and str(val).strip().lower() in JUNK_APPOINTMENT_TYPES


def is_valid_patient(val) -> bool:
    if val is None:
        return False
    s = str(val).strip()
    return not (s == "" or s.lower() == "nan" or _NUMERIC_ONLY_RE.match(s))


def parse_workbook(input_filepath: str) -> list[dict]:
    print(f"   📂 Processing: {os.path.basename(input_filepath)}")
    excel_file = pd.ExcelFile(input_filepath)
    file_rows = []

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
        current_parent = None
        idx = 0

        while idx < len(df):
            row = df.iloc[idx].tolist()
            if len(row) < 3:
                idx += 1; continue
            if pd.isna(row[0]) and pd.isna(row[1]) and pd.isna(row[2]):
                idx += 1; continue

            row_str_0 = str(row[0]).strip()
            row_str_1 = str(row[1]).strip()

            if "APPOINTMENT DATE" in row_str_0 or "Note" == row_str_0:
                idx += 1; continue

            if "/" in row_str_0 and len(row_str_0) <= 10 and "User" not in row_str_0:
                current_parent = {
                    "APPOINTMENT DATE": row[0], "APPOINTMENT TYPE": row[1], "PATIENT": row[2],
                    "CASE": row[3] if len(row) > 3 else None,
                    "CLINIC": row[4] if len(row) > 4 else None,
                    "CALENDAR NAME": row[5] if len(row) > 5 else None,
                }
                idx += 1; continue

            if "User" in row_str_0 or "Updated Date/Time" in row_str_1:
                idx += 1; continue

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
                    "CHANGED FROM":      row[4] if len(row) > 4 else None,
                    "CHANGED TO":        row[5] if len(row) > 5 else None,
                })

            idx += 1

    print(f"   ✅ Extracted {len(file_rows)} event-log rows from {len(excel_file.sheet_names)} clinic tab(s).")
    return file_rows


def clean_raw_workbook(input_filepath: str) -> pd.DataFrame:
    df = pd.DataFrame(parse_workbook(input_filepath))
    if df.empty:
        return df

    before = len(df)
    df = df[~df["APPOINTMENT TYPE"].apply(is_junk_appointment_type)]
    df = df[df["PATIENT"].apply(is_valid_patient)]
    if before != len(df):
        print(f"   🧹 Removed {before - len(df)} junk row(s).")

    before = len(df)
    df = df.drop_duplicates(keep="first")
    if before != len(df):
        print(f"   🧹 Removed {before - len(df)} EXACT duplicate row(s).")

    return df


def main():
    if len(sys.argv) != 3:
        print("Usage: python cleaner.py <raw_input.xlsx> <cleaned_output.csv>")
        sys.exit(1)

    df = clean_raw_workbook(sys.argv[1])
    if df.empty:
        print("⚠️ No rows survived cleaning — nothing written.")
        sys.exit(0)

    df.to_csv(sys.argv[2], index=False, encoding="utf-8-sig")
    print(f"✅ Wrote {len(df)} cleaned row(s) to {sys.argv[2]}")


if __name__ == "__main__":
    main()
