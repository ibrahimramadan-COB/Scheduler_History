"""
scripts/snowflake_writer.py — delete-and-replace per EVENT into
SCHEDULER_HISTORY_DATA2. Now also writes CHANGED_FROM / CHANGED_TO.

Requires the table to have these 2 new columns (run once):
    ALTER TABLE EXTERNAL_DATA.PUBLIC.SCHEDULER_HISTORY_DATA2 ADD COLUMN CHANGED_FROM STRING;
    ALTER TABLE EXTERNAL_DATA.PUBLIC.SCHEDULER_HISTORY_DATA2 ADD COLUMN CHANGED_TO STRING;

Usage: python snowflake_writer.py <cleaned_input.csv>
"""

import os
import sys
import hashlib
import pandas as pd
import snowflake.connector
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.backends import default_backend

TARGET_TABLE = "EXTERNAL_DATA.PUBLIC.SCHEDULER_HISTORY_DATA2"
STAGING_TABLE = "EXTERNAL_DATA.PUBLIC._STAGING_EVENTS"
FIRST_SEEN_LOOKUP_TABLE = "EXTERNAL_DATA.PUBLIC._OLD_FIRST_SEEN"

EVENT_KEY_FIELDS = ["APPOINTMENT DATE", "APPOINTMENT TYPE", "PATIENT", "CASE", "CLINIC", "CALENDAR NAME"]

COLUMN_RENAME = {
    "APPOINTMENT DATE": "APPOINTMENT_DATE", "APPOINTMENT TYPE": "APPOINTMENT_TYPE", "PATIENT": "PATIENT",
    "CASE": "CASE_TITLE", "CLINIC": "CLINIC", "CALENDAR NAME": "CALENDAR_NAME",
    "CREATOR USER": "CREATOR_USER", "CREATED TIMESTAMP": "CREATED_TIMESTAMP", "EVENT ACTION": "EVENT_ACTION",
    "DETAILS": "DETAILS", "CHANGED FROM": "CHANGED_FROM", "CHANGED TO": "CHANGED_TO",
    "SOURCE_FILE": "SOURCE_FILE", "DATA_SOURCE_TAB": "DATA_SOURCE_TAB",
}
DATA_COLUMNS = ["APPOINTMENT_DATE", "APPOINTMENT_TYPE", "PATIENT", "CASE_TITLE", "CLINIC",
                "CALENDAR_NAME", "CREATOR_USER", "CREATED_TIMESTAMP", "EVENT_ACTION",
                "DETAILS", "CHANGED_FROM", "CHANGED_TO", "SOURCE_FILE", "DATA_SOURCE_TAB"]


def get_connection():
    account = os.environ["SNOWFLAKE_ACCOUNT"]
    user = os.environ["SNOWFLAKE_USER"]
    warehouse = os.environ["SNOWFLAKE_WAREHOUSE"]
    database = os.environ.get("SNOWFLAKE_DATABASE", "EXTERNAL_DATA")
    schema = os.environ.get("SNOWFLAKE_SCHEMA", "PUBLIC")
    role = os.environ.get("SNOWFLAKE_ROLE")

    passphrase = os.environ.get("SNOWFLAKE_PRIVATE_KEY_PASSPHRASE") or None
    passphrase_bytes = passphrase.encode() if passphrase else None

    key_file = os.environ.get("SNOWFLAKE_PRIVATE_KEY_FILE")
    if key_file:
        with open(key_file, "rb") as f:
            key_bytes = f.read()
    else:
        key_bytes = os.environ["SNOWFLAKE_PRIVATE_KEY"].encode()

    private_key = serialization.load_pem_private_key(key_bytes, password=passphrase_bytes, backend=default_backend())
    private_key_der = private_key.private_bytes(
        encoding=serialization.Encoding.DER,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    )

    connect_kwargs = dict(user=user, account=account, private_key=private_key_der,
                           warehouse=warehouse, database=database, schema=schema)
    if role:
        connect_kwargs["role"] = role
    return snowflake.connector.connect(**connect_kwargs)


def compute_event_key(row) -> str:
    d = pd.to_datetime(row["APPOINTMENT DATE"], errors="coerce")
    date_str = d.strftime("%Y-%m-%d") if pd.notna(d) else ""
    parts = [date_str] + [str(row[f]).strip() for f in EVENT_KEY_FIELDS[1:]]
    return hashlib.sha256("‖".join(parts).encode("utf-8")).hexdigest()


def prepare_dataframe_from_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()
    df["EVENT_KEY"] = df.apply(compute_event_key, axis=1)
    df["APPOINTMENT DATE"] = pd.to_datetime(df["APPOINTMENT DATE"], errors="coerce").dt.strftime("%Y-%m-%d")
    df["CREATED TIMESTAMP"] = pd.to_datetime(df["CREATED TIMESTAMP"], errors="coerce").dt.strftime("%Y-%m-%d")

    df = df.rename(columns=COLUMN_RENAME)
    for col in DATA_COLUMNS:
        if col not in df.columns:
            df[col] = None
    return df[["EVENT_KEY"] + DATA_COLUMNS]


def prepare_dataframe(csv_path: str) -> pd.DataFrame:
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    return prepare_dataframe_from_df(df)


def load_to_snowflake(conn, df: pd.DataFrame, run_id: str):
    if df.empty:
        print("⚠️ No rows to load.")
        return

    cursor = conn.cursor()
    try:
        cursor.execute(f"CREATE OR REPLACE TEMPORARY TABLE {STAGING_TABLE} LIKE {TARGET_TABLE}")

        cols = ["EVENT_KEY"] + DATA_COLUMNS
        placeholders = ", ".join(["%s"] * len(cols))
        values = df[cols].astype(object).where(pd.notnull(df[cols]), None).values.tolist()
        cursor.executemany(f"INSERT INTO {STAGING_TABLE} ({', '.join(cols)}) VALUES ({placeholders})", values)
        print(f"📤 Staged {len(values)} row(s), {df['EVENT_KEY'].nunique()} distinct event(s).")

        cursor.execute(f"""
            CREATE OR REPLACE TEMPORARY TABLE {FIRST_SEEN_LOOKUP_TABLE} AS
            SELECT EVENT_KEY, MIN(FIRST_SEEN_AT) AS OLD_FIRST_SEEN
            FROM {TARGET_TABLE}
            WHERE EVENT_KEY IN (SELECT DISTINCT EVENT_KEY FROM {STAGING_TABLE})
            GROUP BY EVENT_KEY
        """)
        cursor.execute(f"""
            DELETE FROM {TARGET_TABLE}
            WHERE EVENT_KEY IN (SELECT DISTINCT EVENT_KEY FROM {STAGING_TABLE})
        """)
        print(f"🗑️ Deleted {cursor.rowcount} existing row(s) for events being replaced.")

        insert_cols = ["EVENT_KEY"] + DATA_COLUMNS + ["FIRST_SEEN_AT", "LAST_UPDATED_AT", "SOURCE_RUN_ID"]
        cursor.execute(f"""
            INSERT INTO {TARGET_TABLE} ({', '.join(insert_cols)})
            SELECT
                s.EVENT_KEY, {', '.join('s.' + c for c in DATA_COLUMNS)},
                COALESCE(o.OLD_FIRST_SEEN, CURRENT_TIMESTAMP()) AS FIRST_SEEN_AT,
                CURRENT_TIMESTAMP() AS LAST_UPDATED_AT,
                %s AS SOURCE_RUN_ID
            FROM {STAGING_TABLE} s
            LEFT JOIN {FIRST_SEEN_LOOKUP_TABLE} o ON s.EVENT_KEY = o.EVENT_KEY
        """, (run_id,))
        print(f"✅ Inserted {cursor.rowcount} fresh row(s) into {TARGET_TABLE}.")
    finally:
        cursor.close()


def main():
    raw_file = os.environ.get("RAW_FILE")
    if not raw_file:
        print("Error: RAW_FILE env var not set.")
        sys.exit(1)

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from cleaner import clean_raw_workbook

    cleaned_df = clean_raw_workbook(raw_file)
    if cleaned_df.empty:
        print("⚠️ No rows survived cleaning — nothing to load.")
        return

    run_id = os.environ.get("GITHUB_RUN_ID", "manual_run")
    df = prepare_dataframe_from_df(cleaned_df)
    conn = get_connection()
    try:
        load_to_snowflake(conn, df, run_id)
    finally:
        conn.close()


if __name__ == "__main__":
    main()
