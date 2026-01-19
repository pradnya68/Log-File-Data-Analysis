# v1.07 — Additives Fixed + Progress + Actual Litre + Recipe Usage Charts Working
# ---------------------------------------------------------------------------
# This script parses Hydra dispensing logs from multiple .TXT files.
# It extracts dispensing data (recipe, additives, hydra time, etc.),
# and generates a structured Excel report containing:
#   1️⃣ Dispense_Data (summary of all logs)
#   2️⃣ Individual_Logs (raw parsed data per log file)
#   3️⃣ Usage_Charts (recipe usage frequency + charts)
# ---------------------------------------------------------------------------

import re
import pandas as pd
import glob
import os
from datetime import datetime

# -------------------------------
# Mapping Additive IDs to Names
# -------------------------------
# Each Ad# corresponds to a real chemical additive.
ad_map = {
    "Ad1": "S",
    "Ad2": "PR",
    "Ad3": "T",
    "Ad4": "Brine"
}

# ---------------------------------------------------------------------------
# Function: parse_log(log_file)
# Purpose : Parse a single log file and extract dispensing-related information.
# ---------------------------------------------------------------------------
def parse_log(log_file):
    dispense_data = []   # Stores parsed data entries
    current = {}         # Temporary holder for the current dispense cycle

    # Open the log file safely (ignore encoding errors)
    with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue  # Skip empty lines

            # 🔹 1. Detect start of a dispensing cycle
            start_match = re.search(
                r"Start\s+TankDisp\s+RcpBtIdx=(\d+)\s+(H\d+)\s+Amnt=(\d+)dL",
                line, re.IGNORECASE
            )
            if start_match:
                # Initialize current record
                current = {
                    "LogFile": os.path.basename(log_file),
                    "StartTime": line.split("~")[0].strip() if "~" in line else "",
                    "RecipeIndex": start_match.group(2),  # e.g., H7
                    "Hydra_ms": 0,
                    "Additives": "",
                    "Progress": "",
                    "Actual_Litre": ""
                }

            # 🔹 2. Extract Hydra and Additive timing details
            add_match = re.search(
                r"Hydra=(\d+)ms.*Ad1=(\d+)ms.*Ad2=(\d+)ms.*Ad3=(\d+)ms.*Ad4=(\d+)ms",
                line
            )
            if add_match and current:
                hydra_ms = int(add_match.group(1))
                ad_values = list(map(int, add_match.groups()[1:]))  # Ad1..Ad4 ms
                additives_used = []

                # Only include additives with non-zero timing
                for i, ms in enumerate(ad_values, start=1):
                    if ms > 0:
                        ad_name = ad_map.get(f"Ad{i}", f"Ad{i}")
                        additives_used.append(f"{ad_name}: {ms}ms")

                current["Hydra_ms"] = hydra_ms
                current["Additives"] = ", ".join(additives_used) if additives_used else "None"

            # 🔹 3. Capture dispensing progress (Done/Need)
            prog_match = re.search(r"Disp-Progress\s+Done=(\d+)dL\s+Need=(\d+)dL", line)
            if prog_match and current:
                done = int(prog_match.group(1))
                need = int(prog_match.group(2))
                current["Progress"] = f"{done}/{need} dL"
                current["Actual_Litre"] = f"{done/10:.1f}L"  # Convert dL → L

            # 🔹 4. Detect end of dispensing cycle
            end_match = re.search(r"TankDisp-End\s+Done=(\d+)dL\s+Ret=(\d+)", line)
            if end_match and current:
                # If Actual_Litre was not recorded yet, calculate it
                if not current.get("Actual_Litre"):
                    done = int(end_match.group(1))
                    current["Actual_Litre"] = f"{done/10:.1f}L"

                # Append the completed cycle to the data list
                dispense_data.append(current)
                current = {}  # Reset for next dispense

    return dispense_data


# ---------------------------------------------------------------------------
# STEP 1: Collect all .TXT log files from the current directory
# ---------------------------------------------------------------------------
files = list(set(glob.glob("*.TXT") + glob.glob("*.txt")))
print("Files found:", files)

all_data = []          # Combined data from all logs
individual_data = {}   # Separate data per file

# Parse each log file and store results
for file in files:
    parsed = parse_log(file)
    if parsed:
        df = pd.DataFrame(parsed)
        all_data.extend(parsed)
        individual_data[os.path.basename(file)] = df

# Handle case: No logs found
if not all_data:
    print("⚠️ No dispense data found in logs.")
    exit()

# ---------------------------------------------------------------------------
# STEP 2: Combine all parsed data into one DataFrame
# ---------------------------------------------------------------------------
df = pd.DataFrame(all_data)
df = df[["LogFile", "StartTime", "RecipeIndex", "Hydra_ms", "Additives", "Progress", "Actual_Litre"]]

# ---------------------------------------------------------------------------
# STEP 3: Create unique Excel filename (avoid overwriting)
# ---------------------------------------------------------------------------
today = datetime.today().strftime("%Y-%m-%d")
base_name = f"Combined_Dispense_Log_{today}"
output_file = f"{base_name}.xlsx"
counter = 1
while os.path.exists(output_file):
    output_file = f"{base_name}_{counter}.xlsx"
    counter += 1

# ---------------------------------------------------------------------------
# STEP 4: Generate Excel with 3 sheets using XlsxWriter
# ---------------------------------------------------------------------------
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    workbook = writer.book

    # ============================
    # SHEET 1: Combined Summary
    # ============================
    df.to_excel(writer, sheet_name="Dispense_Data", index=False)
    ws1 = writer.sheets["Dispense_Data"]

    # Define cell formats
    fmt1 = workbook.add_format({"bg_color": "#FFFFFF", "border": 1})
    fmt2 = workbook.add_format({"bg_color": "#F2F2F2", "border": 1})
    hdr_fmt = workbook.add_format({"bold": True, "bg_color": "#DCE6F1", "border": 1, "align": "center"})

    # Freeze header row
    ws1.freeze_panes(1, 0)

    # Write header row with style
    for c, v in enumerate(df.columns.values):
        ws1.write(0, c, v, hdr_fmt)

    # Write data rows with alternating colors
    for r in range(1, len(df)+1):
        f = fmt1 if r % 2 else fmt2
        for c in range(len(df.columns)):
            ws1.write(r, c, df.iloc[r-1, c], f)

    # Enable Excel filters and adjust column widths
    ws1.autofilter(0, 0, len(df), len(df.columns)-1)
    for i, col in enumerate(df.columns):
        ws1.set_column(i, i, max(df[col].astype(str).map(len).max(), len(col)) + 2)

    # ============================
    # SHEET 2: Individual Logs
    # ============================
    ws2 = workbook.add_worksheet("Individual_Logs")
    row = 0
    for file_name, file_df in individual_data.items():
        # Section header for each file
        ws2.write(row, 0, file_name, workbook.add_format({"bold": True, "font_size": 14, "bg_color": "#CFE2F3"}))
        row += 1

        # Write table header
        for c, v in enumerate(file_df.columns.values):
            ws2.write(row, c, v, hdr_fmt)

        # Write file-specific data
        for r in range(len(file_df)):
            f = fmt1 if r % 2 else fmt2
            for c in range(len(file_df.columns)):
                ws2.write(row + 1 + r, c, file_df.iloc[r, c], f)

        # Leave some space before next file
        row += len(file_df) + 3

    # ============================
    # SHEET 3: Recipe Usage Charts
    # ============================
    ws3 = workbook.add_worksheet("Usage_Charts")
    chart_row = 0
    combined_recipe_counts = pd.Series(dtype=int)

    # Generate per-file recipe frequency + chart
    for file_name, file_df in individual_data.items():
        recipe_counts = file_df["RecipeIndex"].value_counts().sort_index()
        combined_recipe_counts = combined_recipe_counts.add(recipe_counts, fill_value=0)

        # Write data for this file
        ws3.write(chart_row, 0, file_name, workbook.add_format({"bold": True, "font_size": 14, "bg_color": "#CFE2F3"}))
        chart_row += 1
        ws3.write_row(chart_row, 0, ["Recipe", "Count"], hdr_fmt)

        # Fill recipe counts
        for i, (recipe, count) in enumerate(recipe_counts.items()):
            ws3.write_row(chart_row + i + 1, 0, [recipe, int(count)], fmt1)

        # Create bar chart for this file
        chart = workbook.add_chart({"type": "column"})
        chart.add_series({
            "name": file_name,
            "categories": ["Usage_Charts", chart_row + 1, 0, chart_row + len(recipe_counts), 0],
            "values": ["Usage_Charts", chart_row + 1, 1, chart_row + len(recipe_counts), 1],
            "data_labels": {"value": True},
        })
        chart.set_title({"name": f"Recipe Usage - {file_name}"})
        chart.set_x_axis({"name": "Recipe Index"})
        chart.set_y_axis({"name": "Count"})
        chart.set_style(10)

        # Insert chart beside data
        ws3.insert_chart(chart_row, 3, chart)
        chart_row += len(recipe_counts) + 15

    # Combined chart across all logs
    ws3.write(chart_row, 0, "All Files Combined", workbook.add_format({"bold": True, "font_size": 14, "bg_color": "#CFE2F3"}))
    chart_row += 1
    ws3.write_row(chart_row, 0, ["Recipe", "Total Count"], hdr_fmt)

    for i, (recipe, count) in enumerate(combined_recipe_counts.sort_index().items()):
        ws3.write_row(chart_row + i + 1, 0, [recipe, int(count)], fmt1)

    # Create combined bar chart
    chart = workbook.add_chart({"type": "column"})
    chart.add_series({
        "name": "All Files Combined",
        "categories": ["Usage_Charts", chart_row + 1, 0, chart_row + len(combined_recipe_counts), 0],
        "values": ["Usage_Charts", chart_row + 1, 1, chart_row + len(combined_recipe_counts), 1],
        "data_labels": {"value": True},
    })
    chart.set_title({"name": "Recipe Usage Across All Files"})
    chart.set_x_axis({"name": "Recipe Index"})
    chart.set_y_axis({"name": "Total Count"})
    chart.set_style(11)

    ws3.insert_chart(chart_row, 3, chart)

# ---------------------------------------------------------------------------
# STEP 5: Print completion summary
# ---------------------------------------------------------------------------
print(f"✅ Excel created successfully with:\n"
      f"   1️⃣ Dispense_Data (Summary)\n"
      f"   2️⃣ Individual_Logs (Raw per file)\n"
      f"   3️⃣ Usage_Charts (Recipe frequency charts)\n"
      f"📊 File saved as: {output_file}")
