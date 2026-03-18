# save as make_serial_list.py
# Usage (from the same folder as the Excel files):
#   pip install pandas openpyxl
#   python make_serial_list.py

import pandas as pd
from pathlib import Path

INPUT_FILE  = "origin1.xlsx"
OUTPUT_FILE = "output1.xlsx"
OUTPUT_SHEET = "Sheet1"   # change if you want a different sheet name

# The exact columns you want in output1.xlsx (from your screenshot)
OUTPUT_HEADERS = [
    "Model", "BMC User", "BMC Password", "BMC MAC (1G)",
    "nic1mac-onb", "nic2mac-onb", "nic3mac-aoc", "nic4mac-aoc",
    "Serial Number", "PO number", "Invoice", "sft-dcms -single", "sft-oob-lic"
]

def find_column_case_insensitive(df, name):
    """Return the actual column name that matches `name` case-insensitively, or None."""
    low = {c.lower(): c for c in df.columns}
    return low.get(name.lower())

def main():
    # --- 1) Read the source workbook (first sheet by default) ---
    if not Path(INPUT_FILE).exists():
        raise FileNotFoundError(f"Can't find {INPUT_FILE} in {Path.cwd()}")

    df = pd.read_excel(INPUT_FILE, engine="openpyxl")

    # --- 2) Locate the SERIALNUM column robustly (case-insensitive) ---
    serial_col = find_column_case_insensitive(df, "SERIALNUM")
    if not serial_col:
        raise KeyError("Could not find a 'SERIALNUM' column in origin1.xlsx")

    # --- 3) Clean up & get unique serials in the order they first appear ---
    # Keep as strings to avoid Excel scientific notation for long IDs
    serials = (
        df[serial_col]
        .dropna()
        .astype(str)
        .str.strip()
    )

    # pandas.unique preserves first-occurrence order
    unique_serials = pd.unique(serials)

    # --- 4) Build the output dataframe with your exact headers ---
    out = pd.DataFrame({"Serial Number": unique_serials})
    # Reindex to your full header list (others remain blank)
    out = out.reindex(columns=OUTPUT_HEADERS)

    # --- 5) Write output1.xlsx ---
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as xw:
        out.to_excel(xw, index=False, sheet_name=OUTPUT_SHEET)

    print(f"Found {len(unique_serials)} unique serial(s).")
    print(f"Wrote {OUTPUT_FILE} with column 'Serial Number' filled.")

if __name__ == "__main__":
    main()