#!/bin/bash

# Build script for macOS
# This script creates a standalone macOS application

echo "Building Excel Combiner for macOS..."

# Check if required packages are installed
echo "Installing required packages..."
pip3 install pandas openpyxl xlrd pyinstaller --user

# Create the macOS app bundle
echo "Creating macOS application bundle..."
pyinstaller excel_combiner.spec --clean --noconfirm

# Check if build was successful
if [ -d "dist/ExcelCombiner.app" ]; then
    echo "✅ macOS build completed successfully!"
    echo "📁 Application created at: dist/ExcelCombiner.app"
    echo ""
    echo "To run the application:"
    echo "  Double-click on ExcelCombiner.app in the dist folder"
    echo "  Or run: open dist/ExcelCombiner.app"
    echo ""
    echo "To create a distributable package:"
    echo "  You can compress the ExcelCombiner.app folder to create a .zip file"
    echo "  Or use tools like 'create-dmg' to create a .dmg installer"
else
    echo "❌ Build failed. Check the output above for errors."
    exit 1
fi