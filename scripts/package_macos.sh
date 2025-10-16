#!/bin/bash

# Package the macOS application for distribution
# This script creates a distributable .zip file

echo "Packaging Excel Combiner for macOS distribution..."

# Check if the app exists
if [ ! -d "dist/ExcelCombiner.app" ]; then
    echo "❌ ExcelCombiner.app not found. Please build the application first."
    echo "Run: ./build_macos.sh"
    exit 1
fi

# Create a distribution folder
mkdir -p distribution
rm -rf distribution/ExcelCombiner_macOS

# Copy the app and documentation
cp -r dist/ExcelCombiner.app distribution/
cp README_GUI.md distribution/README.md

# Create the zip file
cd distribution
zip -r "ExcelCombiner_macOS_v1.0.0.zip" ExcelCombiner.app README.md
cd ..

echo "✅ Distribution package created!"
echo "📦 File: distribution/ExcelCombiner_macOS_v1.0.0.zip"
echo ""
echo "Contents:"
echo "  - ExcelCombiner.app (macOS application)"
echo "  - README.md (user guide)"
echo ""
echo "Ready for distribution!"