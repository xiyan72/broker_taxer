"""
Script to match tax_stock_name in lastest_buy_date.ods to 证券名称 in
stock_name_code_exchange_inventory.xlsx and fill the latest_buy column.

Matching strategy (improved):
1. The short name (证券名称) must be a SUBSTRING of the full name (tax_stock_name),
   OR characters of short name appear in order within the full name.
2. Score = (# chars of short_name found consecutively in full_name) / len(short_name)
3. Threshold = 0.6 (to avoid false positives)
4. Duplicates in ODS (same company, same value) are deduplicated.
5. When multiple different values match the same xlsx row, join with ' / '.

Manual overrides are applied for edge cases that fuzzy matching gets wrong.
"""

import pandas as pd
import re
from difflib import SequenceMatcher

ODS_FILE = r'C:\Users\1\Documents\myCode\broker_taxer\lastest_buy_date.ods'
XLSX_FILE = r'C:\Users\1\Documents\myCode\broker_taxer\stock_name_code_exchange_inventory.xlsx'
OUTPUT_FILE = XLSX_FILE  # overwrite in place

# Manual overrides: (partial_tax_name_match) -> xlsx_证券名称
# Use these when automatic matching would produce wrong results.
# Key = substring that uniquely identifies the tax_stock_name
# Value = exact 证券名称 in xlsx, or None to explicitly skip
MANUAL_OVERRIDES = {
    '中国国际海运集装箱': '中远海运港口',  # COSCO container -> COSCO ports
    '中国中铁': None,                        # 中国中铁 is NOT in xlsx — skip
}


def normalize(s):
    """Strip whitespace and full-width spaces."""
    return str(s).strip().replace(' ', '').replace('　', '')


def score_match(short_name, full_name):
    """
    Score how well a short name (证券名称) matches a full company name (tax_stock_name).
    Returns a float in [0, 1].
    - 1.0 if short_name is a direct substring of full_name
    - Otherwise use SequenceMatcher ratio on the best alignment
    """
    s = normalize(short_name)
    f = normalize(full_name)
    if not s or not f or f == 'nan':
        return 0.0

    # Clean 'A' suffix from short names like '万  科Ａ'
    s_clean = re.sub(r'[ＡA]$', '', s)

    if s in f or (s_clean and s_clean in f):
        return 1.0

    # Use SequenceMatcher for partial overlap
    ratio = SequenceMatcher(None, s, f).ratio()
    # Also check if most chars of s appear in f (in order)
    matched = 0
    fi = 0
    for ch in s:
        while fi < len(f) and f[fi] != ch:
            fi += 1
        if fi < len(f):
            matched += 1
            fi += 1
    order_score = matched / len(s)

    return max(ratio, order_score)


def find_best_match(tax_name, xlsx_df):
    """
    Check manual overrides first, then find best scoring xlsx row.
    Returns (xlsx_index or None, matched_name or None, score).
    """
    # Check manual overrides
    for key, override_name in MANUAL_OVERRIDES.items():
        if key in tax_name:
            if override_name is None:
                return None, None, 0.0
            # Find the xlsx row with this 证券名称
            matches = xlsx_df.index[xlsx_df['证券名称'] == override_name].tolist()
            if matches:
                return matches[0], override_name, 1.0
            else:
                print(f"    WARNING: override target '{override_name}' not found in xlsx!")
                return None, None, 0.0

    # Automatic matching
    best_idx = None
    best_score = 0.0
    best_name = None
    for idx, row in xlsx_df.iterrows():
        short = row['证券名称']
        sc = score_match(short, tax_name)
        if sc > best_score:
            best_score = sc
            best_idx = idx
            best_name = short

    return best_idx, best_name, best_score


def main():
    SCORE_THRESHOLD = 0.65

    # ---- Load files ----
    print("Loading files...")
    df_ods = pd.read_excel(ODS_FILE, engine='odf')
    df_xlsx = pd.read_excel(XLSX_FILE, engine='openpyxl')

    print(f"ODS rows: {len(df_ods)}, XLSX rows: {len(df_xlsx)}")

    # Ensure latest_buy column exists and is clear
    df_xlsx['latest_buy'] = None

    # ---- Match ODS rows to XLSX rows ----
    # xlsx_idx -> set of (stock_time_number values) to avoid duplicates
    match_results = {}   # xlsx_idx -> list of unique stock_time_number strings

    unmatched = []

    print("\n--- Matching ODS rows to XLSX rows ---")
    for _, row in df_ods.iterrows():
        tax_name = str(row.get('tax_stock_name', '')).strip()
        time_number = str(row.get('stock_time_number', '')).strip()
        row_no = row.get('no', '')

        # Skip rows that have no meaningful tax_name
        if not tax_name or tax_name == 'nan' or len(tax_name) < 3:
            print(f"  ODS row {str(row_no):>6}: SKIP (empty/split row)")
            continue

        # Skip rows with no time_number value
        if not time_number or time_number == 'nan':
            print(f"  ODS row {str(row_no):>6}: SKIP (no stock_time_number)")
            continue

        best_idx, best_name, best_score = find_best_match(tax_name, df_xlsx)

        if best_idx is not None and best_score >= SCORE_THRESHOLD:
            print(f"  ODS row {str(row_no):>6}: '{tax_name}' -> '{best_name}' (score={best_score:.2f}) | '{time_number}'")
            if best_idx not in match_results:
                match_results[best_idx] = []
            # Avoid exact duplicates
            if time_number not in match_results[best_idx]:
                match_results[best_idx].append(time_number)
        else:
            score_info = f"best score={best_score:.2f} ('{best_name}')" if best_name else "no candidates"
            print(f"  ODS row {str(row_no):>6}: '{tax_name}' -> NO MATCH [{score_info}]")
            unmatched.append((row_no, tax_name, time_number))

    # ---- Populate latest_buy in xlsx ----
    print("\n--- Updating XLSX latest_buy column ---")
    for xlsx_idx, values in match_results.items():
        combined = ' / '.join(values)
        df_xlsx.at[xlsx_idx, 'latest_buy'] = combined
        print(f"  XLSX row {xlsx_idx:>3}: '{df_xlsx.at[xlsx_idx, '证券名称']}' -> '{combined}'")

    # ---- Report unmatched ----
    if unmatched:
        print(f"\n--- Unmatched ODS rows ({len(unmatched)}) ---")
        for row_no, tax_name, time_number in unmatched:
            print(f"  Row {row_no}: '{tax_name}' | '{time_number}'")

    # ---- Save output ----
    print(f"\nSaving to {OUTPUT_FILE} ...")
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df_xlsx.to_excel(writer, index=False)
    print("Done!")

    # ---- Summary ----
    filled = df_xlsx[df_xlsx['latest_buy'].notna()]
    empty = df_xlsx[df_xlsx['latest_buy'].isna()]
    print(f"\n=== SUMMARY ===")
    print(f"Rows updated: {len(filled)} / {len(df_xlsx)}")
    print(f"\nFilled rows:")
    for _, r in filled.iterrows():
        print(f"  {r['证券名称']:20s}: {r['latest_buy']}")
    print(f"\nNot filled (no ODS match):")
    for _, r in empty.iterrows():
        print(f"  {r['证券名称']}")


if __name__ == '__main__':
    main()
