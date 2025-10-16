# Excel Combiner GUI - User Guide

A comprehensive guide to using the Excel Combiner GUI application for combining multiple Excel files into a single spreadsheet.

## 🚀 Quick Start

### Installation Options

#### Option 1: Download Pre-built Executable
1. **Windows**: Download `ExcelCombiner.exe` from the GitHub releases or `dist_win11/` folder
2. **macOS**: Download `ExcelCombiner.app` from the `dist/` folder
3. **Run**: Double-click the executable - no installation required!

#### Option 2: Run from Source
```bash
# Clone or download the project
# Install dependencies
pip install -r requirements.txt

# Run the GUI
python excel_combiner_gui.py
```

## 🖥️ Application Interface

### Main Window Components

#### 1. **Folder Selection Section**
- **"Browse Folder" Button**: Click to select the folder containing your Excel files
- **Selected Path Display**: Shows the currently selected folder path
- **Status**: Displays "No folder selected" until you choose a folder

#### 2. **File List Display**
- **File Count**: Shows total number of Excel files found (e.g., "Found 4 Excel files")
- **File Names**: Lists all detected `.xlsx`, `.xls` files in the selected folder
- **Auto-refresh**: Updates automatically when you select a new folder

#### 3. **Output Options**
- **Output Filename Field**: Enter desired name for combined file
- **Default**: Pre-filled with `combined_excel_files.xlsx`
- **Auto-extension**: `.xlsx` added automatically if not provided

#### 4. **Action Buttons**
- **"Combine Excel Files" Button**: Starts the combination process
- **Status**: Disabled until folder is selected and files are found
- **Progress Feedback**: Shows "Processing..." during operation

#### 5. **Progress Display**
- **Progress Bar**: Visual indicator of combination progress
- **Percentage**: Shows completion percentage (0-100%)
- **Current File**: Displays which file is currently being processed

#### 6. **Results Section**
- **Success Message**: "✅ Successfully combined X files into filename.xlsx"
- **Total Rows**: Shows final row count in combined file
- **Processing Time**: Displays how long the operation took
- **Error Display**: Shows any errors encountered during processing

## 📋 Step-by-Step Usage

### Step 1: Launch the Application
- **Windows**: Double-click `ExcelCombiner.exe`
- **macOS**: Double-click `ExcelCombiner.app` (may need to approve in Security settings)
- **Source**: Run `python excel_combiner_gui.py`

### Step 2: Select Your Excel Files Folder
1. Click the **"Browse Folder"** button
2. Navigate to the folder containing your Excel files
3. Click **"Select Folder"** or **"Choose"**
4. The application will automatically scan for Excel files

### Step 3: Review Detected Files
- Check the file list to ensure all desired files are detected
- The application finds files with extensions: `.xlsx`, `.xls`
- Files are processed in alphabetical order

### Step 4: Configure Output
1. **Output Name**: Enter desired filename (default: `combined_excel_files.xlsx`)
2. **Location**: File will be saved in the same folder as your source files
3. **Format**: Output is always Excel format (`.xlsx`)

### Step 5: Combine Files
1. Click **"Combine Excel Files"** button
2. **Progress Tracking**: Watch the progress bar and current file indicator
3. **Wait**: Process completes automatically (usually takes a few seconds)
4. **Confirmation**: Success message appears when complete

### Step 6: Access Results
- **Output File**: Located in your source folder with specified name
- **Open**: Double-click the combined file to open in Excel
- **Verify**: Check that all data from source files is included

## 🔧 Advanced Features

### Source File Tracking
The application automatically adds a "Source File" column to track which original file each row came from:
- **Column Header**: "Source File"
- **Content**: Original filename for each row
- **Grouping**: Rows from the same file are grouped together
- **Unique Display**: Each filename appears only once per group (not repeated for every row)

### Data Processing Details
- **Header Handling**: Uses headers from the first Excel file encountered
- **Sheet Selection**: Processes the first sheet in each Excel file
- **Row Preservation**: Maintains all data rows from all source files
- **Column Alignment**: Automatically aligns columns across different files

### File Format Support
- **Primary**: `.xlsx` (Excel 2007+) - Recommended format
- **Legacy**: `.xls` (Excel 97-2003) - Also supported
- **Mixed**: Can combine both formats in same operation
- **Output**: Always saves as `.xlsx` format

## 🐛 Troubleshooting

### Common Issues and Solutions

#### 1. **"No Excel files found" Message**
**Problem**: Folder selected but no files detected
**Solutions**:
- Verify files have `.xlsx` or `.xls` extensions
- Check files aren't in subfolders (only scans selected folder)
- Ensure files aren't corrupted or password-protected

#### 2. **Application Won't Start**
**Problem**: Executable doesn't launch
**Solutions**:
- **Windows**: Right-click → "Run as administrator" or approve SmartScreen warning
- **macOS**: System Preferences → Security & Privacy → "Open Anyway"
- **Source**: Verify Python 3.8+ and dependencies installed

#### 3. **"Permission Denied" Error**
**Problem**: Can't save output file
**Solutions**:
- Ensure output file isn't already open in Excel
- Check write permissions for target folder
- Close any Excel applications that might lock the file

#### 4. **"Memory Error" During Processing**
**Problem**: Large files cause memory issues
**Solutions**:
- Process smaller batches of files
- Close other applications to free memory
- Use 64-bit Python if processing very large datasets

#### 5. **Missing Data in Output**
**Problem**: Some data doesn't appear in combined file
**Solutions**:
- Verify all source files have consistent column headers
- Check for hidden sheets (only first sheet is processed)
- Ensure source files aren't corrupted

#### 6. **Slow Processing**
**Problem**: Combination takes very long
**Solutions**:
- **Normal**: Large files (>1MB each) naturally take longer
- **Optimization**: Close other applications during processing
- **Hardware**: More RAM and faster storage improve performance

### Performance Guidelines
- **Small Files** (<1MB each): Near-instant processing
- **Medium Files** (1-10MB each): Few seconds per file
- **Large Files** (>10MB each): May take 30+ seconds per file
- **Memory Usage**: Approximately 2-5x the total size of input files

## 💡 Tips and Best Practices

### Data Preparation
1. **Consistent Headers**: Ensure all Excel files have same column headers
2. **Clean Data**: Remove empty rows/columns from source files for best results
3. **File Names**: Use descriptive filenames - they appear in the "Source File" column
4. **Backup**: Keep copies of original files before combining

### Workflow Optimization
1. **Organize Files**: Place all files to combine in a dedicated folder
2. **Preview First**: Open a few source files to verify structure before combining
3. **Test Small**: Try with 2-3 files first to verify results
4. **Batch Processing**: For large datasets, consider combining in smaller groups

### Output Management
1. **Descriptive Names**: Use meaningful output filenames (e.g., `Q4_sales_data_combined.xlsx`)
2. **Version Control**: Include dates in filenames for tracking (e.g., `data_combined_2024-01-15.xlsx`)
3. **Verification**: Always open and spot-check the combined file after processing

## 🔍 Technical Specifications

### System Requirements
- **Windows**: 10/11 (64-bit recommended)
- **macOS**: 10.14 Mojave or later
- **Memory**: 4GB RAM minimum, 8GB+ recommended for large files
- **Storage**: 100MB for application, plus 2-5x source file sizes for processing

### File Specifications
- **Input Formats**: `.xlsx`, `.xls`
- **Output Format**: `.xlsx` (Excel 2007+)
- **Maximum File Size**: Limited by available system memory
- **Maximum Files**: No limit (memory dependent)

### Data Limitations
- **Rows**: Excel's limit of 1,048,576 rows per sheet
- **Columns**: Excel's limit of 16,384 columns per sheet
- **Cell Content**: Excel's standard text/numeric limitations

## 📞 Support and Resources

### Getting Help
- **Documentation**: Check this user guide and build instructions
- **GitHub Issues**: Report bugs or request features
- **Source Code**: Available for customization and extension

### Project Links
- **Repository**: https://github.com/legithubgris/excel-combiner-gui
- **Releases**: Latest executables and changelogs
- **Documentation**: Complete guides and technical details

The Excel Combiner GUI is designed to be intuitive and reliable for combining Excel files quickly and efficiently. This guide covers all standard usage scenarios - for advanced customization, refer to the source code and technical documentation.