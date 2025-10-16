# Excel Combiner Tool

A cross-platform desktop application for combining multiple Excel files into a single file with source tracking and an intuitive graphical user interface.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-macOS%20%7C%20Windows-lightgrey.svg)
![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)

## 🚀 Features

- **🎨 Full-Row Highlighting**: NEW! Preserve and extend highlight colors across entire rows for enhanced visual scanning
- **Cross-platform**: Runs on macOS and Windows
- **User-friendly GUI**: Simple point-and-click interface built with tkinter
- **Real-time feedback**: Progress updates and detailed logging
- **Smart file handling**: Automatically excludes output files from processing
- **Source tracking**: Shows source filename only once per file group
- **Error handling**: Comprehensive error reporting and validation
- **Threading**: Non-blocking processing that keeps the GUI responsive
- **Automated builds**: GitHub Actions CI/CD for Windows executables

## 📁 Project Structure

```
📁 Excel Combiner Tool/
├── 📄 README.md                    # This file
├── 📄 requirements.txt             # Python dependencies
├── 🐍 excel_combiner_gui.py        # Main GUI application
├── 🐍 combine_excel_files.py       # Command-line version
├── ⚙️  excel_combiner.spec         # macOS PyInstaller config
├── ⚙️  excel_combiner_windows.spec # Windows PyInstaller config
├── 📁 dist/                        # macOS executables
│   ├── ExcelCombiner               # CLI executable
│   └── ExcelCombiner.app/          # GUI application bundle
├── 📁 dist_win11/                  # Windows executables
│   └── ExcelCombiner.exe           # GUI executable
├── 📁 scripts/                     # Build automation
│   ├── build_macos.sh              # macOS build script
│   ├── build_windows.bat           # Windows build script
│   └── package_macos.sh            # macOS packaging script
├── 📁 docs/                        # Documentation
│   ├── BUILD_INSTRUCTIONS.md       # Detailed build guide
│   └── README_GUI.md               # GUI-specific documentation
└── 📁 sample_data/                 # Example Excel files
    ├── part1.xlsx
    ├── part2.xlsx
    ├── part3.xlsx
    └── part4.xlsx
```

## 🎯 Quick Start

### Option 1: Use Pre-built Executables

#### macOS Users
1. Double-click `dist/ExcelCombiner.app`
2. If macOS shows a security warning, go to System Preferences > Security & Privacy and click "Open Anyway"

#### Windows Users
1. Double-click `dist_win11/ExcelCombiner.exe`
2. Windows may show a SmartScreen warning - click "More info" then "Run anyway"

### Option 2: Run from Source
```bash
# Install dependencies
pip install -r requirements.txt

# Run the GUI application
python excel_combiner_gui.py

# Or run the command-line version
python combine_excel_files.py /path/to/excel/files output_filename.xlsx
```

## 🛠️ How to Use

1. **Launch the application**
   - macOS: Double-click `ExcelCombiner.app`
   - Windows: Double-click `ExcelCombiner.exe`
   - From source: Run `python excel_combiner_gui.py`

2. **Select source folder**
   - Click "Browse" next to "Source Folder"
   - Navigate to the folder containing your Excel files
   - Click "Select Folder"

3. **Specify output filename** (optional)
   - Default: "combined_excel_files.xlsx"
   - Change if you want a different name

4. **Combine files**
   - Click "Combine Excel Files"
   - Watch the progress and log for updates
   - Success message appears when complete

## 📊 How It Works

The application processes Excel files by:

1. **Reading** columns A, B, and C from each Excel file (typically: Filename, Transcription, Status)
2. **🎨 Detecting** highlighted rows and preserving formatting information
3. **Combining** all data into a single Excel file with enhanced full-row highlighting
4. **Adding** source filename in column D, shown only once at the start of each file's data
5. **Extending** highlight colors across entire rows (25+ columns) for improved visual scanning
6. **Preserving** the header row only once at the top
7. **Processing** files in alphabetical order
8. **Excluding** previous output files to prevent duplicates

### ✨ Full-Row Highlighting Feature

NEW! The application now enhances row visibility:
- **Color Detection**: Automatically identifies cells with background colors or highlighting
- **Full-Row Extension**: Extends detected colors across the entire row width for improved visual scanning
- **Format Preservation**: Maintains original font styles, colors, and other formatting
- **Visual Continuity**: Makes it easier to track highlighted data across wide spreadsheets

### Example Output Structure:
```
Row 1:  [Header]  [Header]  [Header]  [Source_File]
Row 2:  [Data]    [Data]    [Data]    file1.xlsx
Row 3:  [Data]    [Data]    [Data]    
Row 4:  [Data]    [Data]    [Data]    
Row 32: [Data]    [Data]    [Data]    file2.xlsx
Row 33: [Data]    [Data]    [Data]    
Row 66: [Data]    [Data]    [Data]    file3.xlsx
```

## 🔧 Development

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Setup Development Environment
```bash
# Clone or download the project
# Navigate to the project directory

# Create virtual environment (recommended)
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# or
.venv\Scripts\activate     # Windows

# Install dependencies
pip install -r requirements.txt
```

### Building Executables

#### macOS
```bash
./scripts/build_macos.sh
```

#### Windows
```bash
scripts\build_windows.bat
```

Or use the automated GitHub Actions workflow by pushing to the main branch.

## 🤖 Automated Builds

This project uses GitHub Actions for automated Windows builds:

- ✅ **Automatic builds** on every push to main branch
- ✅ **Artifact uploads** with 30-day retention
- ✅ **Build verification** to ensure executable works
- ✅ **Manual trigger** support via workflow_dispatch

Access builds at: [GitHub Actions](https://github.com/legithubgris/excel-combiner-gui/actions)

## 📋 Supported File Formats

- **Input**: `.xlsx` and `.xls` files
- **Output**: `.xlsx` format

## 🐛 Troubleshooting

### Common Issues

1. **"Cannot open Excel file"**
   - Ensure files are not open in Excel
   - Check file permissions
   - Verify file format (.xlsx or .xls)

2. **Application won't start**
   - macOS: Check Security & Privacy settings
   - Windows: Allow through SmartScreen filter
   - From source: Ensure Python dependencies are installed

3. **Memory issues with large files**
   - Close other applications
   - Process files in smaller batches
   - Ensure adequate disk space

### Getting Help

1. Check the application log for specific error messages
2. Ensure your Excel files follow the expected format (columns A, B, C)
3. Verify you have read/write permissions for the target folders
4. For development issues, see `docs/BUILD_INSTRUCTIONS.md`

## 📝 License

This project is open source. Feel free to use, modify, and distribute.

## 🤝 Contributing

Contributions welcome! Please feel free to submit pull requests or open issues.

## 📧 Support

For issues or questions:
1. Check the troubleshooting section above
2. Review the detailed documentation in the `docs/` folder
3. Create an issue on the GitHub repository

---

**Version**: 2.0  
**Last Updated**: October 10, 2025  
**Compatibility**: macOS 10.14+, Windows 10+

## Installation

1. Make sure you have Python 3.6+ installed
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install pandas openpyxl xlrd
```

## Usage

### Basic Usage

```bash
# Combine Excel files in the current directory
python combine_excel_files.py

# Combine Excel files in a specific folder
python combine_excel_files.py /path/to/excel/files/

# Specify a custom output filename
python combine_excel_files.py /path/to/excel/files/ -o my_combined_file.xlsx
```

### Examples

```bash
# Process files in the current directory
python combine_excel_files.py

# Process files in a specific folder
python combine_excel_files.py "/Users/username/Documents/ExcelFiles/"

# Process files and save with custom name
python combine_excel_files.py . -o "consolidated_data.xlsx"
```

## Output

The script creates a new Excel file with:
- Column A: Filename (from original files)
- Column B: Transcription (from original files) 
- Column C: Status (from original files)
- Column D: Source_File (name of the Excel file, shown only once at the start of each file's data group)

### Example Output Structure:
```
Row 1:  [Data] [Data] [Data] file1.xlsx
Row 2:  [Data] [Data] [Data] 
Row 3:  [Data] [Data] [Data] 
Row 31: [Data] [Data] [Data] file2.xlsx
Row 32: [Data] [Data] [Data] 
Row 66: [Data] [Data] [Data] file3.xlsx
```

## Notes

- The script automatically sorts input files alphabetically for consistent processing order
- Header row is included only once at the top of the combined file
- Empty files are skipped
- Error handling is included for corrupted or unreadable files
- Progress information is displayed during processing