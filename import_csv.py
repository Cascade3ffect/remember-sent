# import_csv.py

import pandas as pd
import sqlite3
from datetime import datetime
from pathlib import Path

imports_dir = Path("imports")

csv_files = list(imports_dir.glob("SentEmailAnalysis*_results*.csv"))

if not csv_files:
    raise FileNotFoundError(f"No matching CSV files found in {imports_dir}")

csv_file = max(
    csv_files,
    key=lambda f: f.stat().st_mtime
)

print(f"Importing: {csv_file}...")

df = pd.read_csv(csv_file)
df["imported_at"] = datetime.now()

print("Converting to email.db...")
conn = sqlite3.connect("email.db")
df.to_sql(
    "conversations",
    conn,
    if_exists="replace",
    index=False
)
conn.close()

print(f"Imported {len(df)} rows")