#!/usr/bin/env python3
import argparse
import difflib
import os
import sys
import pandas as pd

def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("data", help="the txt file or directory")
    parser.add_argument("-r", "--reference", help="the excel file (optional)")
    parser.add_argument("-o", "--output", help="output file")
    return parser.parse_args()

def parse_xls(filename):
    fmt_list = []
    df = pd.read_excel(filename, dtype=str).fillna("")
    start_row = 0
    for idx, row in df.iterrows():
        if "STARTING" in row.values and "FIELD" in row.values:
            start_row = idx
            break

    df.columns = df.iloc[start_row]
    df = df.iloc[start_row+1:].reset_index(drop=True)

    df = df[df["STARTING"].astype(str).str.strip() != ""].reset_index(drop=True)

    starting_col = [i for i, col in enumerate(df.columns) if col == "STARTING"][0]
    field_col = [i for i, col in enumerate(df.columns) if col == "FIELD"][0]

    for _, row in df.iterrows():
        offset_val = row.iloc[starting_col]
        if offset_val == "BYTE":
            continue
        offset = int(offset_val) - 1
        field_name = row.iloc[field_col]
        fmt_list.append({"offset": offset, "field": field_name})

    return fmt_list

def parse_txt(filename, fmt):
    field_names = [entry["field"] for entry in fmt]

    colspecs = []
    prev_offset = fmt[0]["offset"]
    for entry in fmt[1:]:
        colspecs.append((prev_offset, entry["offset"]))
        prev_offset = entry["offset"]
    colspecs.append((prev_offset, -1))

    df = pd.read_fwf(filename, colspecs=colspecs, names=field_names, dtype=str)
    return df

def find_best_xls_match(txt_file):
    dir_path = os.path.dirname(os.path.abspath(txt_file))
    txt_basename = os.path.splitext(os.path.basename(txt_file))[0]
    txt_basename = txt_basename.replace(" extract data format", "")

    xls_files = [f for f in os.listdir(dir_path) if f.lower().endswith((".xls", ".xlsx"))]
    if not xls_files:
        raise FileNotFoundError(f"No Excel files found in {dir_path}")

    best_match = None
    best_ratio = 0
    for xls in xls_files:
        xls_base = os.path.splitext(xls)[0]
        ratio = difflib.SequenceMatcher(None, txt_basename.lower(), xls_base.lower()).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_match = xls

    if best_match is None:
        raise FileNotFoundError("No matching Excel file found.")

    return os.path.join(dir_path, best_match)

def process_file(txt_file, ref_file=None, output_file=None):
    if ref_file is None:
        ref_file = find_best_xls_match(txt_file)
        # choice = input(f"Use reference file {ref_file}? [y/N]: ")
        choice = 'y'
        if choice and choice[0].lower() != 'y':
            return

    fmt = parse_xls(ref_file)
    df = parse_txt(txt_file, fmt)

    if output_file is None:
        base_name, _ = os.path.splitext(os.path.basename(txt_file))
        output_file = base_name + ".csv"

    df.to_csv(output_file, index=False)
    print(output_file)

if __name__ == "__main__":
    args = parse_args()

    if not os.path.isdir(args.data):
        process_file(args.data, ref_file = args.reference, output_file = args.output)
        sys.exit(0)

    for file in os.listdir(args.data):
        if file.lower().endswith(".txt"):
            txt_path = os.path.join(args.data, file)
            try:
                process_file(txt_path)
            except Exception as e:
                print(f"Error processing {txt_path}: {e}")
