# Excel Combiner GUI Application

A cross-platform desktop application for combining multiple Excel files into a single file with an intuitive graphical user interface.

## Features

- **Cross-platform**: Runs on macOS and Windows
- **User-friendly GUI**: Simple point-and-click interface
- **Real-time feedback**: Progress updates and detailed logging
- **Smart file handling**: Automatically excludes output files from processing
- **Clean output format**: Source filenames shown only once per file group
- **Error handling**: Comprehensive error reporting and validation
- **Threading**: Non-blocking processing that keeps the GUI responsive

## System Requirements

### macOS:
- macOS 10.14 (Mojave) or later
- No additional dependencies required for the standalone app

### Windows:
- Windows 10 or later
- No additional dependencies required for the standalone executable

## Download & Installation

### macOS:
1. Download `ExcelCombiner.app` (or the .dmg installer if available)
2. Move the app to your Applications folder
3. Right-click and select "Open" the first time (to bypass Gatekeeper)
4. For subsequent uses, double-click to launch

### Windows:
1. Download `ExcelCombiner.exe`
2. Save it to a folder of your choice
3. Double-click to launch
4. Windows may show a security warning - click "More info" then "Run anyway"

## How to Use

1. **Launch the application**
   - macOS: Double-click ExcelCombiner.app
   - Windows: Double-click ExcelCombiner.exe

2. **Select source folder**
   - Click "Browse" next to "Source Folder"
   - Navigate to the folder containing your Excel files
   - Click "Select Folder"

3. **Specify output filename** (optional)
   - The default is "combined_excel_files.xlsx"
   - Change if you want a different name

4. **Combine files**
   - Click "Combine Excel Files"
   - Watch the progress and log for updates
   - A success message will appear when complete

5. **Find your combined file**
   - The output file will be saved in the same folder as your source files
   - Check the log for the exact location

## What the Application Does

The application:
- Reads columns A, B, and C from each Excel file (Filename, Transcription, Status)
- Combines all data into a single Excel file
- Adds the source filename in column D, but only once at the start of each file's data
- Preserves the header row only once at the top
- Processes files in alphabetical order
- Automatically excludes previous output files to prevent duplicates

### Example Output Structure:
```
Row 1:  [Header] [Header] [Header] [Source_File]
Row 2:  [Data]   [Data]   [Data]   file1.xlsx
Row 3:  [Data]   [Data]   [Data]   
Row 4:  [Data]   [Data]   [Data]   
Row 32: [Data]   [Data]   [Data]   file2.xlsx
Row 33: [Data]   [Data]   [Data]   
Row 66: [Data]   [Data]   [Data]   file3.xlsx
```

## Supported File Formats

- **Input**: .xlsx and .xls files
- **Output**: .xlsx format

## Troubleshooting

### macOS Issues:

**"App can't be opened because it's from an unidentified developer"**
- Right-click the app and select "Open"
- Click "Open" in the dialog that appears
- Or run: `xattr -cr /path/to/ExcelCombiner.app`

**App won't launch**
- Make sure you're running macOS 10.14 or later
- Try running from Terminal: `open ExcelCombiner.app`

### Windows Issues:

**Windows Defender warning**
- Click "More info" then "Run anyway"
- This is common with PyInstaller executables and is a false positive
- You can add the executable to Windows Defender's exclusion list

**Antivirus blocking the executable**
- Add the executable to your antivirus exclusion list
- This is a common false positive with packaged Python applications

### General Issues:

**"No Excel files found" message**
- Make sure your folder contains .xlsx or .xls files
- Check that the files aren't corrupted
- Ensure you have read permissions for the folder

**Processing fails**
- Check the log for specific error messages
- Ensure Excel files have the expected column structure (A, B, C)
- Make sure you have write permissions in the target folder

**Large file processing is slow**
- This is normal for files with many rows
- The progress bar will continue to animate while processing
- Don't close the application while processing

## Features of the GUI

- **Folder Browser**: Easy navigation to your Excel files
- **Output Filename**: Customizable output file name
- **Progress Bar**: Visual feedback during processing
- **Detailed Log**: Real-time updates and error messages
- **Clear Log**: Button to clear the log display
- **Status Bar**: Current operation status
- **Responsive Design**: Window can be resized as needed

## Technical Information

- **Built with**: Python 3.13, tkinter, pandas, openpyxl
- **Packaged with**: PyInstaller 6.15.0
- **Architecture**: Native ARM64 (Apple Silicon) and x64 support
- **Threading**: Background processing to prevent GUI freezing

## Version History

### Version 1.0.0
- Initial release
- Cross-platform GUI application
- Excel file combination with source tracking
- Real-time progress and logging
- Standalone executables for macOS and Windows

## Support

For issues or questions:
1. Check this README for common solutions
2. Look at the application log for specific error messages
3. Ensure your Excel files follow the expected format (columns A, B, C)

## Development

If you want to run from source code:
```bash
# Install dependencies
pip install pandas openpyxl xlrd

# Run the application
python excel_combiner_gui.py
```

For building from source, see `BUILD_INSTRUCTIONS.md`.