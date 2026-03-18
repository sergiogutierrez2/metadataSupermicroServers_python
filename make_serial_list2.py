# make_serial_list2.py
# Requirements: pandas, openpyxl
# Files expected in the SAME folder as this script:
#   - origin1.xlsx  (source with columns: ORDERNUM, SERVERPARTNO, SERIALNUM, Assembly Date, SUB-ITEM, SUB-SERIAL)
#   - output1.xlsx  (target with a column "Serial Number" and your other headers)
#
# What it does:
# - Reads all serials from output1.xlsx ("Serial Number" column, in order).
# - For each serial, searches origin1.xlsx rows with that SERIALNUM.
# - Fills Model, BMC Password, BMC MAC (1G), nic1mac-onb, nic2mac-onb, nic3mac-aoc, nic4mac-aoc.
# - Optionally fills BMC User if SUB-ITEM == "NUM-DEFUSR" exists.
# - Writes back to output1.xlsx (if it's locked, saves as output1_<timestamp>.xlsx).

from pathlib import Path
from datetime import datetime
import pandas as pd

ORIGIN_FILE = "origin1.xlsx"
OUTPUT_FILE = "output1.xlsx"
OUTPUT_SHEET = None  # None = first sheet

# Columns to fill in output1.xlsx
TARGET_COLS = [
    "Model",
    "BMC User",
    "BMC Password",
    "BMC MAC (1G)",
    "nic1mac-onb",
    "nic2mac-onb",
    "nic3mac-aoc",
    "nic4mac-aoc",
]

def _find_col(df, name):
    """Case-insensitive, space-trim tolerant column resolver."""
    m = {c.strip().lower(): c for c in df.columns}
    return m.get(name.strip().lower())

def _as_str_series(s):
    return s.dropna().astype(str).str.strip()

def _first(series):
    vals = _as_str_series(series)
    return vals.iloc[0] if not vals.empty else ""

def _first_n(series, n):
    vals = _as_str_series(series).tolist()
    # pad with empty strings to always return length n
    vals += [""] * (n - len(vals))
    return vals[:n]

def main():
    here = Path.cwd()

    # --- Read origin ---
    if not Path(ORIGIN_FILE).exists():
        raise FileNotFoundError(f"Can't find {ORIGIN_FILE} in {here}")

    origin = pd.read_excel(ORIGIN_FILE, engine="openpyxl")

    # Resolve origin columns robustly
    col_serial   = _find_col(origin, "SERIALNUM")
    col_serverpn = _find_col(origin, "SERVERPARTNO")
    col_subitem  = _find_col(origin, "SUB-ITEM")
    col_subserial= _find_col(origin, "SUB-SERIAL")
    if not all([col_serial, col_serverpn, col_subitem, col_subserial]):
        raise KeyError("origin1.xlsx must contain columns SERIALNUM, SERVERPARTNO, SUB-ITEM, SUB-SERIAL (any case).")

    # Normalize key columns used for matching
    origin[col_serial] = origin[col_serial].astype(str).str.strip()
    origin[col_subitem] = origin[col_subitem].astype(str).str.strip()

    # --- Read output (target) ---
    if not Path(OUTPUT_FILE).exists():
        raise FileNotFoundError(f"Can't find {OUTPUT_FILE} in {here}")

    out = pd.read_excel(OUTPUT_FILE, sheet_name=0, engine="openpyxl")
    col_serial_out = _find_col(out, "Serial Number")
    if not col_serial_out:
        raise KeyError("output1.xlsx must have a 'Serial Number' column (any case).")

    # Ensure target columns exist; if not, add them blank
    for c in TARGET_COLS:
        if c not in out.columns:
            out[c] = ""

    # Work on a copy of the serial list to preserve order
    serials = out[col_serial_out].astype(str).str.strip()

    # --- Build a lookup per serial ---
    filled_rows = []
    for sn in serials:
        if not sn or sn.lower() == "nan":
            # blank serial row; append placeholders
            filled_rows.append({
                "Model": "",
                "BMC User": "",
                "BMC Password": "",
                "BMC MAC (1G)": "",
                "nic1mac-onb": "",
                "nic2mac-onb": "",
                "nic3mac-aoc": "",
                "nic4mac-aoc": "",
            })
            continue

        grp = origin[origin[col_serial] == sn]

        # Model from SERVERPARTNO (first occurrence)
        model = _first(grp[col_serverpn])

        # BMC User (optional): SUB-ITEM == NUM-DEFUSR (if present)
        bmc_user = _first(grp.loc[grp[col_subitem] == "NUM-DEFUSR", col_subserial])

        # BMC Password: SUB-ITEM == NUM-DEFPWD
        bmc_pwd = _first(grp.loc[grp[col_subitem] == "NUM-DEFPWD", col_subserial])

        # BMC MAC (1G): SUB-ITEM == MAC-IPMI-ADDRESS
        bmc_mac = _first(grp.loc[grp[col_subitem] == "MAC-IPMI-ADDRESS", col_subserial])

        # Onboard NICs: up to two MAC-ADDRESS
        nic_onb_1, nic_onb_2 = _first_n(grp.loc[grp[col_subitem] == "MAC-ADDRESS", col_subserial], 2)

        # AOC NICs: up to two MAC-AOC-ADDRESS
        nic_aoc_3, nic_aoc_4 = _first_n(grp.loc[grp[col_subitem] == "MAC-AOC-ADDRESS", col_subserial], 2)

        filled_rows.append({
            "Model": model,
            "BMC User": bmc_user,                # blank if none found
            "BMC Password": bmc_pwd,
            "BMC MAC (1G)": bmc_mac,
            "nic1mac-onb": nic_onb_1,
            "nic2mac-onb": nic_onb_2,
            "nic3mac-aoc": nic_aoc_3,
            "nic4mac-aoc": nic_aoc_4,
        })

    # Create a DataFrame aligned to out rows
    filled_df = pd.DataFrame(filled_rows, index=out.index)

    # Update/overwrite only the target columns
    for c in TARGET_COLS:
        out[c] = filled_df[c].astype(str)

    # --- Write back (handle locked file gracefully) ---
    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as xw:
            out.to_excel(xw, index=False, sheet_name=OUTPUT_SHEET or "Sheet1")
        print(f"Updated '{OUTPUT_FILE}' successfully.")
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = OUTPUT_FILE.replace(".xlsx", f"_{ts}.xlsx")
        with pd.ExcelWriter(alt, engine="openpyxl") as xw:
            out.to_excel(xw, index=False, sheet_name=OUTPUT_SHEET or "Sheet1")
        print(f"'{OUTPUT_FILE}' was open. Saved as '{alt}' instead.")

if __name__ == "__main__":
    main()