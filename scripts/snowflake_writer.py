"""
Takes the cleaned (all-event-types) DataFrame from cleaner.py, computes
a deterministic ROW_KEY per row (hash of the 9 identity columns), stages
it as CSV in Snowflake, and MERGEs it into SCHEDULER_HISTORY_DATA —
inserting genuinely new rows and refreshing existing ones with whatever
this run's scrape found, in one atomic statement.

No Group/Group 2 join happens here anymore — that's handled by a view
in Snowflake (V_APPOINTMENTS_ENRICHED, joining against
SCHEDULER_HISTORY_NAMES) so department attribution always reflects the
CURRENT Names mapping, not whatever it was at scrape time.

Auth: key-pair (RSA), same pattern as your existing Snowflake notebooks
— no password, no MFA/Duo prompt, safe for an unattended pipeline.
    SNOWFLAKE_PRIVATE_KEY            — PEM content (GitHub secret)
    SNOWFLAKE_PRIVATE_KEY_FILE       — path to a .p8 file (local runs)
    SNOWFLAKE_PRIVATE_KEY_PASSPHRASE — optional, only if the key is encrypted
    SNOWFLAKE_ACCOUNT, SNOWFLAKE_USER, SNOWFLAKE_WAREHOUSE,
    SNOWFLAKE_DATABASE, SNOWFLAKE_SCHEMA, SNOWFLAKE_ROLE
"""

import os
import sys
import csv
import hashlib
import tempfile
from datetime import datetime

import pandas as pd
import snowflake.connector
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.backends import default_backend

TABLE_NAME = "SCHEDULER_HISTORY_DATA"
STAGE_NAME = "SCHEDULER_HISTORY_STAGE"

# Cleaner.py's column names -> Snowflake column names. CASE and DETAILS
# stay similar, but CASE is reserved in SQL, hence CASE_TITLE.
COLUMN_RENAME = {
    "APPOINTMENT DATE": "APPOINTMENT_DATE",
    "APPOINTMENT TYPE": "APPOINTMENT_TYPE",
    "PATIENT": "PATIENT",
    "CASE": "CASE_TITLE",
    "CLINIC": "CLINIC",
    "CALENDAR NAME": "CALENDAR_NAME",
    "CREATOR USER": "CREATOR_USER",
    "CREATED TIMESTAMP": "CREATED_TIMESTAMP",
    "EVENT ACTION": "EVENT_ACTION",
    "DETAILS": "DETAILS",
    "SOURCE_FILE": "SOURCE_FILE",
    "DATA_SOURCE_TAB": "DATA_SOURCE_TAB",
}

# The 9-column identity used everywhere else in this project — this is
# what ROW_KEY is a hash of.
IDENTITY_COLUMNS = [
    "APPOINTMENT_DATE", "APPOINTMENT_TYPE", "PATIENT", "CASE_TITLE",
    "CLINIC", "CALENDAR_NAME", "CREATOR_USER", "CREATED_TIMESTAMP", "EVENT_ACTION",
]

ALL_COLUMNS = IDENTITY_COLUMNS + ["DETAILS", "SOURCE_FILE", "DATA_SOURCE_TAB"]


def get_connection():
    account = os.environ.get("SNOWFLAKE_ACCOUNT")
    user = os.environ.get("SNOWFLAKE_USER")
    warehouse = os.environ.get("SNOWFLAKE_WAREHOUSE")
    database = os.environ.get("SNOWFLAKE_DATABASE", "EXTERNAL_DATA")
    schema = os.environ.get("SNOWFLAKE_SCHEMA", "PUBLIC")
    role = os.environ.get("SNOWFLAKE_ROLE")

    missing = [n for n, v in [("SNOWFLAKE_ACCOUNT", account), ("SNOWFLAKE_USER", user),
                               ("SNOWFLAKE_WAREHOUSE", warehouse)] if not v]
    if missing:
        raise RuntimeError(f"Missing required env var(s): {', '.join(missing)}")

    key_file = os.environ.get("SNOWFLAKE_PRIVATE_KEY_FILE")
    key_content = os.environ.get("SNOWFLAKE_PRIVATE_KEY")
    passphrase = os.environ.get("SNOWFLAKE_PRIVATE_KEY_PASSPHRASE")
    passphrase_bytes = passphrase.encode() if passphrase else None

    if key_file:
        with open(key_file, "rb") as f:
            key_bytes = f.read()
    elif key_content:
        key_bytes = key_content.encode()
    else:
        raise RuntimeError("Set either SNOWFLAKE_PRIVATE_KEY (PEM content) or SNOWFLAKE_PRIVATE_KEY_FILE (path).")

    private_key = serialization.load_pem_private_key(key_bytes, password=passphrase_bytes, backend=default_backend())
    private_key_der = private_key.private_bytes(
        encoding=serialization.Encoding.DER,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    )

    connect_kwargs = dict(
        user=user, account=account, private_key=private_key_der,
        warehouse=warehouse, database=database, schema=schema,
    )
    if role:
        connect_kwargs["role"] = role

    return snowflake.connector.connect(**connect_kwargs)


def build_row_key(df: pd.DataFrame) -> pd.Series:
    def _key_for_row(row):
        parts = []
        for col in IDENTITY_COLUMNS:
            v = row[col]
            parts.append("" if pd.isna(v) else str(v).strip())
        raw = "‖".join(parts)
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()
    return df.apply(_key_for_row, axis=1)


def prepare_dataframe(cleaned_df: pd.DataFrame) -> pd.DataFrame:
    df = cleaned_df.rename(columns=COLUMN_RENAME).copy()
    for col in ALL_COLUMNS:
        if col not in df.columns:
            df[col] = None

    # Date-only for both date fields — matches the DATE column types and
    # the identity semantics used throughout this project.
    df["APPOINTMENT_DATE"] = pd.to_datetime(df["APPOINTMENT_DATE"], errors="coerce").dt.strftime("%Y-%m-%d")
    df["CREATED_TIMESTAMP"] = pd.to_datetime(df["CREATED_TIMESTAMP"], errors="coerce").dt.strftime("%Y-%m-%d")

    df = df[ALL_COLUMNS]
    df["ROW_KEY"] = build_row_key(df)

    before = len(df)
    df = df.drop_duplicates(subset="ROW_KEY", keep="last")
    if len(df) != before:
        print(f"   🧹 Removed {before - len(df)} duplicate row(s) within this batch "
              f"(MERGE requires unique keys in the source).")

    df["SOURCE_RUN_ID"] = os.environ.get("GITHUB_RUN_ID", datetime.now().strftime("local_%Y%m%d_%H%M%S"))
    return df


def ensure_stage(cursor):
    cursor.execute(f"CREATE STAGE IF NOT EXISTS {STAGE_NAME}")


def load_to_snowflake(df: pd.DataFrame) -> dict:
    if df.empty:
        print("   ℹ️ No rows to load.")
        return {"inserted": 0, "updated": 0}

    conn = get_connection()
    cursor = conn.cursor()
    try:
        ensure_stage(cursor)

        tmp_path = os.path.join(tempfile.gettempdir(), "scheduler_history_batch.csv")
        stage_columns = ["ROW_KEY"] + ALL_COLUMNS + ["SOURCE_RUN_ID"]
        df[stage_columns].to_csv(tmp_path, index=False, na_rep="", encoding="utf-8-sig",
                                  quotechar='"', quoting=csv.QUOTE_ALL)

        staging_table = f"{TABLE_NAME}_BATCH"
        cursor.execute(f"""
            CREATE OR REPLACE TEMPORARY TABLE {staging_table} LIKE {TABLE_NAME}
        """)
        # TEMPORARY TABLE LIKE copies structure including defaults; that's fine,
        # we'll supply every column explicitly on COPY INTO / MERGE anyway.

        put_path = tmp_path.replace("\\", "/")
        cursor.execute(f"PUT file://{put_path} @{STAGE_NAME} AUTO_COMPRESS=FALSE OVERWRITE=TRUE")

        col_list = ", ".join(stage_columns)
        cursor.execute(f"""
            COPY INTO {staging_table} ({col_list})
            FROM @{STAGE_NAME}/scheduler_history_batch.csv
            FILE_FORMAT = (TYPE = CSV FIELD_OPTIONALLY_ENCLOSED_BY = '"' SKIP_HEADER = 1
                            NULL_IF = ('') ENCODING = 'UTF8')
        """)

        set_clause = ", ".join(f"target.{c} = source.{c}" for c in ALL_COLUMNS)
        insert_cols = ["ROW_KEY"] + ALL_COLUMNS + ["SOURCE_RUN_ID", "FIRST_SEEN_AT", "LAST_UPDATED_AT"]
        insert_values = ["source.ROW_KEY"] + [f"source.{c}" for c in ALL_COLUMNS] + \
                         ["source.SOURCE_RUN_ID", "CURRENT_TIMESTAMP()", "CURRENT_TIMESTAMP()"]

        cursor.execute(f"""
            MERGE INTO {TABLE_NAME} AS target
            USING {staging_table} AS source
            ON target.ROW_KEY = source.ROW_KEY
            WHEN MATCHED THEN UPDATE SET
                {set_clause},
                target.SOURCE_RUN_ID = source.SOURCE_RUN_ID,
                target.LAST_UPDATED_AT = CURRENT_TIMESTAMP()
            WHEN NOT MATCHED THEN INSERT ({", ".join(insert_cols)})
            VALUES ({", ".join(insert_values)})
        """)

        result = cursor.fetchone()
        inserted, updated = (result[0], result[1]) if result else (0, 0)

        os.remove(tmp_path)
        return {"inserted": inserted, "updated": updated}
    finally:
        cursor.close()
        conn.close()


def write_to_snowflake(cleaned_df: pd.DataFrame, start_date: str, end_date: str) -> dict:
    df = prepare_dataframe(cleaned_df)
    print(f"   📤 Loading {len(df)} row(s) into {TABLE_NAME}...")
    result = load_to_snowflake(df)
    print(f"   ✅ MERGE complete: {result['inserted']} inserted, {result['updated']} updated.")
    return {"rows_scraped": len(df), **result}


if __name__ == "__main__":
    import sys
    from cleaner import clean_raw_workbook

    raw_file = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("RAW_FILE")
    start_date = os.environ.get("START_DATE", "")
    end_date = os.environ.get("END_DATE", "")
    if not raw_file:
        print("❌ No raw file path given (arg or RAW_FILE env var).")
        sys.exit(1)

    cleaned_df = clean_raw_workbook(raw_file)
    result = write_to_snowflake(cleaned_df, start_date, end_date)
    print(f"\n✅ Done. Scraped {result['rows_scraped']} rows: "
          f"{result['inserted']} inserted, {result['updated']} updated.")
