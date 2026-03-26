"""
Batch runner — scans both sheets from the same workbook sequentially.
Edit the CONFIGURATION section below, then run:
    python scan_both_sheets.py
"""

import subprocess
import sys
import time

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION — Edit these values to match your file
# ═══════════════════════════════════════════════════════════════════════════════

INPUT_FILE = "your_file.xlsx"          # ← your input xlsx file
USERNAME_COLUMN = "A"                   # ← column letter with usernames

# Sheet 1: 700-1000 usernames
SHEET_1_NAME = "700-1000"              # ← exact sheet name
SHEET_1_OUTPUT = "results_700_1000.xlsx"

# Sheet 2: >1000 usernames
SHEET_2_NAME = ">1000"                 # ← exact sheet name
SHEET_2_OUTPUT = "results_1000_plus.xlsx"

HEADLESS = True                         # True = invisible browser, False = watch it work

# ═══════════════════════════════════════════════════════════════════════════════


def run_sheet(sheet_name, output_file):
    cmd = [
        sys.executable, "twitter_scanner.py",
        INPUT_FILE,
        "--sheet", sheet_name,
        "--column", USERNAME_COLUMN,
        "--output", output_file,
    ]
    if HEADLESS:
        cmd.append("--headless")

    print(f"\n{'='*60}")
    print(f"  SCANNING SHEET: {sheet_name}")
    print(f"  Output → {output_file}")
    print(f"{'='*60}\n")

    result = subprocess.run(cmd)
    return result.returncode


def main():
    print("Twitter/X Profile Scanner — Batch Mode")
    print(f"Input file: {INPUT_FILE}\n")

    # Scan Sheet 1
    code1 = run_sheet(SHEET_1_NAME, SHEET_1_OUTPUT)
    if code1 != 0:
        print(f"\nWarning: Sheet 1 scan exited with code {code1}")

    # Brief pause between sheets
    print(f"\n  Pausing 60s before scanning next sheet...\n")
    time.sleep(60)

    # Scan Sheet 2
    code2 = run_sheet(SHEET_2_NAME, SHEET_2_OUTPUT)
    if code2 != 0:
        print(f"\nWarning: Sheet 2 scan exited with code {code2}")

    print(f"\n{'='*60}")
    print(f"  ALL DONE!")
    print(f"  Results: {SHEET_1_OUTPUT}, {SHEET_2_OUTPUT}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
