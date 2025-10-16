#!/usr/bin/env python3
"""
Specific debug for formatting issue
"""

import sys
sys.path.append('/Users/gr4yf1r3/Library/CloudStorage/OneDrive-Nuance/audioMover/_migration/walgreens_excelPlayground/ReDooV2/en_transcriptions_locationprompt_Tuned_9.15.2025_For_ScriptSplits_part1-4')

from excel_combiner_gui import ExcelCombinerGUI
import tkinter as tk

# Let's test just the formatting extraction
app = ExcelCombinerGUI(tk.Tk())

# Test on the file we know has formatting
test_file = "/Users/gr4yf1r3/Library/CloudStorage/OneDrive-Nuance/audioMover/_migration/walgreens_excelPlayground/ReDooV2/en_transcriptions_locationprompt_Tuned_9.15.2025_For_ScriptSplits_part1-4/sample_data/en_transcriptions_locationprompt_Tuned_9.15.2025_For_ScriptSplits_part4.xlsx"

print("Testing formatting extraction...")
df, filename, row_formats = app.read_excel_data_with_formatting(test_file)

print(f"File: {filename}")
print(f"DataFrame shape: {df.shape}")
print(f"Row formats found: {len(row_formats) if row_formats else 0}")

if row_formats:
    print("First few formatting entries:")
    for i, (row_num, format_data) in enumerate(list(row_formats.items())[:5]):
        print(f"  Row {row_num}: {format_data}")
else:
    print("No formatting found!")