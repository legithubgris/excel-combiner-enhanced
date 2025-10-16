#!/usr/bin/env python3
"""
Simple test of the GUI logic without the actual GUI
"""

import sys
import os

# Add the current directory to the path so we can import the GUI module
sys.path.append('/Users/gr4yf1r3/Library/CloudStorage/OneDrive-Nuance/audioMover/_migration/walgreens_excelPlayground/ReDooV2/en_transcriptions_locationprompt_Tuned_9.15.2025_For_ScriptSplits_part1-4')

# Import the combiner class
from excel_combiner_gui import ExcelCombinerGUI
import tkinter as tk

class TestCombiner:
    def __init__(self):
        # Create a minimal GUI instance for testing
        root = tk.Tk()
        root.withdraw()  # Hide the window
        self.app = ExcelCombinerGUI(root)
        
        # Set test parameters
        self.app.folder_path.set("/Users/gr4yf1r3/Library/CloudStorage/OneDrive-Nuance/audioMover/_migration/walgreens_excelPlayground/ReDooV2/en_transcriptions_locationprompt_Tuned_9.15.2025_For_ScriptSplits_part1-4/sample_data")
        self.app.output_filename.set("gui_test_combined.xlsx")
    
    def test_combine(self):
        """Test the combine functionality."""
        print("Testing GUI combiner functionality...")
        result = self.app.combine_excel_files()
        return result

def main():
    tester = TestCombiner()
    success = tester.test_combine()
    
    if success:
        print("\n✅ GUI test completed successfully!")
    else:
        print("\n❌ GUI test failed!")

if __name__ == "__main__":
    main()