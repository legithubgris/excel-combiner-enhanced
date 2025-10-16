# Building Excel Combiner Executables

This comprehensive guide covers building standalone executables for both macOS and Windows, including automated builds via GitHub Actions.

## 📋 Prerequisites

### Common Requirements
- Python 3.8 or higher
- All project dependencies installed
- PyInstaller for executable creation

### 🔧 Install Dependencies
```bash
# Install project dependencies
pip install -r requirements.txt

# Install PyInstaller for building executables
pip install pyinstaller
```

## 🍎 macOS Build Process

### Method 1: Automated Build Script
```bash
# Make the build script executable
chmod +x scripts/build_macos.sh

# Run the automated build
./scripts/build_macos.sh
```

### Method 2: Manual Build
```bash
# Build using the macOS spec file
pyinstaller excel_combiner.spec --clean --noconfirm

# Optional: Create app bundle
./scripts/package_macos.sh
```

### 📦 macOS Output
- **Command-line executable**: `dist/ExcelCombiner`
- **GUI app bundle**: `dist/ExcelCombiner.app`

### macOS Build Details
- Creates a native macOS application bundle
- Includes all Python dependencies
- No Python installation required on target machines
- App bundle can be distributed via drag-and-drop

## 🪟 Windows Build Process

### Method 1: Native Windows Build
If building on Windows:
```cmd
scripts\build_windows.bat
```

### Method 2: GitHub Actions (Recommended)
For automated cross-platform Windows builds:

1. **Push to repository**: Any push to `main` branch triggers build
2. **Manual trigger**: Use GitHub Actions "workflow_dispatch" 
3. **Download artifact**: Get `ExcelCombiner-Windows.zip` from Actions tab

#### Accessing GitHub Actions Builds:
- **Repository**: https://github.com/legithubgris/excel-combiner-gui
- **Actions Tab**: https://github.com/legithubgris/excel-combiner-gui/actions
- **Download**: Click on successful run → Download "ExcelCombiner-Windows" artifact

### 📦 Windows Output
- **GUI executable**: `dist_win11/ExcelCombiner.exe` (local) or artifact download
- **Size**: ~35MB (includes all dependencies)
- **Compatibility**: Windows 10/11

## ⚙️ Build Configuration

### PyInstaller Spec Files

#### 🍎 macOS Configuration (`excel_combiner.spec`)
```python
# Key features:
- App bundle creation with proper structure
- Includes tkinter and GUI dependencies  
- Optimized for macOS distribution
- Code signing preparation
```

#### 🪟 Windows Configuration (`excel_combiner_windows.spec`)
```python
# Key features:
- Single-file executable creation
- All dependencies bundled
- Windowed application (no console)
- Optimized for Windows distribution
```

### Hidden Imports Configuration
Both spec files include essential hidden imports:
- `pandas` (Excel processing)
- `openpyxl` (Excel file support)
- `xlrd` (Legacy Excel support)
- `tkinter` components (GUI framework)

## 🤖 Automated Build System

### GitHub Actions Workflow
- **Triggers**: Push to main, pull requests, manual dispatch
- **Environment**: Windows Server 2022, Python 3.11
- **Steps**:
  1. Checkout code
  2. Setup Python environment
  3. Install dependencies (pandas, openpyxl, xlrd, PyInstaller)
  4. Build executable with PyInstaller
  5. Test executable creation
  6. Upload artifact (30-day retention)

### Accessing Automated Builds
```bash
# Local download via MCP tools (as demonstrated)
curl -L [artifact_url] -o ExcelCombiner-Windows.zip
unzip ExcelCombiner-Windows.zip
```

## 🐛 Troubleshooting

### Common Build Issues

#### 1. **Missing Dependencies**
```bash
# Solution: Ensure all packages installed
pip install -r requirements.txt
pip install pyinstaller
```

#### 2. **Import Errors During Build**
```python
# Add to spec file's hiddenimports:
hiddenimports=[
    'pandas',
    'openpyxl', 
    'xlrd',
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog'
]
```

#### 3. **Large Executable Size**
- **Normal**: 30-50MB typical for pandas applications
- **Optimization**: Use `--exclude-module` for unused packages
- **Trade-off**: Size vs. compatibility

#### 4. **Platform-Specific Issues**

**macOS Security Warnings**
```bash
# User solution:
System Preferences > Security & Privacy > "Open Anyway"

# Developer solution:
codesign -s "Developer ID" ExcelCombiner.app
```

**Windows SmartScreen Warnings**
```cmd
# User solution:
"More info" > "Run anyway"

# Developer solution:
signtool sign /f certificate.p12 ExcelCombiner.exe
```

#### 5. **GitHub Actions Build Failures**
- Check Actions tab for detailed logs
- Common issues: dependency conflicts, spec file errors
- Solution: Review build logs and update configurations

## 📦 Distribution Guide

### 🍎 macOS Distribution
1. **Test**: Verify on different macOS versions (10.14+)
2. **Package**: Use `.app` bundle for easy installation
3. **Security**: Consider notarization for Gatekeeper approval
4. **Distribution**: GitHub releases, direct download, or App Store

### 🪟 Windows Distribution  
1. **Test**: Verify on Windows 10/11
2. **Package**: Single `.exe` file, no installation needed
3. **Security**: Consider code signing for SmartScreen approval
4. **Distribution**: GitHub releases, direct download, or Microsoft Store

### 🌐 GitHub Releases Workflow
```bash
# Create tagged release
git tag v2.0
git push origin v2.0

# Automated builds attach to releases
# Download from: github.com/user/repo/releases
```

## 📝 Build Scripts Reference

### `scripts/build_macos.sh`
- Automated macOS build process
- Uses `excel_combiner.spec` configuration
- Creates both CLI and GUI versions
- Handles cleanup and organization

### `scripts/build_windows.bat`
- Windows batch file for local builds
- Uses `excel_combiner_windows.spec` configuration
- Creates single executable file
- Includes error handling

### `scripts/package_macos.sh`
- macOS app bundle creation
- Proper `.app` structure setup
- Icon and metadata inclusion
- Distribution preparation

## 🔄 Continuous Integration

The project includes full CI/CD pipeline:
- **Trigger**: Every push to main branch
- **Build**: Automated Windows executable creation
- **Test**: Executable verification and basic functionality
- **Artifact**: Downloadable Windows executable with 30-day retention
- **Monitoring**: Build status and failure notifications

This ensures every code change produces a tested, distributable Windows executable automatically.