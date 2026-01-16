import os
import pandas as pd
from pathlib import Path
from datetime import datetime
import logging
import numpy as np

# ----------------------------
# LOGGING CONFIG
# ----------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# ----------------------------
# READ FILE (CSV or EXCEL)
# ----------------------------
def read_file(file_path):
    logging.info(f"Reading file: {file_path}")
    if str(file_path).endswith(".csv"):
        return pd.read_csv(file_path)
    elif str(file_path).endswith(".xlsx"):
        return pd.read_excel(file_path)
    else:
        raise ValueError("Only CSV and Excel files are supported")

# ----------------------------
# COMPARE FILES BASED ON KEY COLUMN
# ----------------------------
def compare_files(file1, file2, key_column):
    df1 = read_file(file1)
    df2 = read_file(file2)

    # Get file names without extension
    file1_name = Path(file1).stem
    file2_name = Path(file2).stem

    # Validate key column
    if key_column not in df1.columns or key_column not in df2.columns:
        raise ValueError(f"Key column '{key_column}' must exist in both files")

    df1 = df1.set_index(key_column)
    df2 = df2.set_index(key_column)

    all_keys = sorted(set(df1.index).union(set(df2.index)))
    all_columns = sorted(set(df1.columns).union(set(df2.columns)))

    logging.info(f"Total keys to compare: {len(all_keys)}")
    logging.info(f"Total columns to compare: {len(all_columns)}")

    summary_rows = []
    detailed_rows = []
    missing_rows = []

    # Determine numeric columns
    numeric_cols = [col for col in all_columns if
                    np.issubdtype(df1[col].dtype, np.number) or
                    np.issubdtype(df2[col].dtype, np.number)]

    for key in all_keys:
        # Missing key handling
        if key not in df1.index:
            missing_rows.append({"key": key, "present_in": f"{file2_name}_only"})
            continue
        if key not in df2.index:
            missing_rows.append({"key": key, "present_in": f"{file1_name}_only"})
            continue

        mismatches = 0
        for col in all_columns:
            val1 = df1.at[key, col] if col in df1.columns else None
            val2 = df2.at[key, col] if col in df2.columns else None

            if val1 != val2:
                mismatches += 1
                row = {
                    "key": key,
                    "column_name": col,
                    file1_name: val1,
                    file2_name: val2
                }
                # If numeric, add difference
                if col in numeric_cols:
                    try:
                        row["difference"] = float(val2) - float(val1)
                    except:
                        row["difference"] = None
                detailed_rows.append(row)

        total_cols = len(all_columns)
        match_pct = round(((total_cols - mismatches) / total_cols) * 100, 2)

        summary_rows.append({
            "key": key,
            "total_columns": total_cols,
            "matched_columns": total_cols - mismatches,
            "mismatched_columns": mismatches,
            "match_percentage": match_pct,
            "status": "✅ MATCH" if mismatches == 0 else "❌ MISMATCH"
        })

    return (
        pd.DataFrame(summary_rows),
        pd.DataFrame(detailed_rows),
        pd.DataFrame(missing_rows)
    )

# ----------------------------
# SAVE FORMATTED EXCEL REPORT
# ----------------------------
def save_report(summary_df, detail_df, missing_df, output_file):
    logging.info(f"Saving report to: {output_file}")
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

        # Write sheets
        summary_df.sort_values("match_percentage").to_excel(writer, sheet_name="SUMMARY", index=False)
        detail_df.to_excel(writer, sheet_name="DETAILED_MISMATCHES", index=False)
        missing_df.to_excel(writer, sheet_name="MISSING_KEYS", index=False)

        # Formats
        header_fmt = workbook.add_format({"bold": True, "border": 1, "align": "center"})
        green_fmt = workbook.add_format({"bg_color": "#C6EFCE"})
        red_fmt = workbook.add_format({"bg_color": "#FFC7CE"})

        # ---------------- SUMMARY FORMAT ----------------
        summary_ws = writer.sheets["SUMMARY"]
        summary_ws.freeze_panes(1, 0)
        for col_num, col_name in enumerate(summary_df.columns):
            summary_ws.write(0, col_num, col_name, header_fmt)
            summary_ws.set_column(col_num, col_num, 18)

        summary_ws.conditional_format(
            f"F2:F{len(summary_df)+1}",
            {"type": "text", "criteria": "containing", "value": "MATCH", "format": green_fmt}
        )
        summary_ws.conditional_format(
            f"F2:F{len(summary_df)+1}",
            {"type": "text", "criteria": "containing", "value": "MISMATCH", "format": red_fmt}
        )

        # ---------------- DETAIL FORMAT ----------------
        detail_ws = writer.sheets["DETAILED_MISMATCHES"]
        detail_ws.freeze_panes(1, 0)
        detail_ws.set_column(0, len(detail_df.columns)-1, 25)

        # ---------------- MISSING FORMAT ----------------
        missing_ws = writer.sheets["MISSING_KEYS"]
        missing_ws.freeze_panes(1, 0)
        missing_ws.set_column(0, len(missing_df.columns)-1, 25)

    logging.info("Comparison report saved successfully!")

# ----------------------------
# MAIN
# ----------------------------
if __name__ == "__main__":
    project_dir = Path(__file__).resolve().parent

    # Input files (from Downloads)
    downloads = os.path.expanduser("~/Downloads")
    file1 = Path(downloads) / "cricket_file_100_1.xlsx"
    file2 = Path(downloads) / "cricket_file_100_2.xlsx"

    # Key column for comparison
    key_column = "player_id"

    logging.info("Starting file comparison...")
    summary, details, missing = compare_files(file1, file2, key_column)

    # Save report in script directory with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = project_dir / f"comparison_result_{timestamp}.xlsx"

    save_report(summary, details, missing, output_file)
    logging.info("Comparison Completed!")
