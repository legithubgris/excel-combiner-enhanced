#!/usr/bin/env python3
"""
Debug version to see what's happening
"""

import sys
sys.path.append('/Users/gr4yf1r3/Library/CloudStorage/OneDrive-Nuance/audioMover/_migration/walgreens_excelPlayground/ReDooV2/en_transcriptions_locationprompt_Tuned_9.15.2025_For_ScriptSplits_part1-4')

from excel_combiner_gui import ExcelCombinerGUI
import tkinter as tk

class TestCombinerWithLogs:
    def __init__(self):
        root = tk.Tk()
        root.withdraw()
        
        self.app = ExcelCombinerGUI(root)
        self.app.folder_path.set("/Users/gr4yf1r3/Library/CloudStorage/OneDrive-Nuance/audioMover/_migration/walgreens_excelPlayground/ReDooV2/en_transcriptions_locationprompt_Tuned_9.15.2025_For_ScriptSplits_part1-4/sample_data")
        self.app.output_filename.set("debug_combined.xlsx")
        
        # Override log_message to print to console
        original_log = self.app.log_message
        def console_log(message):
            print(f"LOG: {message}")
            original_log(message)
        self.app.log_message = console_log
    
    def test_combine(self):
        print("Testing GUI combiner with debug logs...")
        result = self.app.combine_excel_files()
        return result

def main():
    tester = TestCombinerWithLogs()
    success = tester.test_combine()
    
    if success:
        print("\n✅ Debug test completed!")
    else:
        print("\n❌ Debug test failed!")

if __name__ == "__main__":
    main()