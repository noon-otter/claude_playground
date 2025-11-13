#!/bin/bash

# Usage: ./install_addin.sh path/to/manifest.xml

set -e

# Excel for Mac sideload directory
TARGET_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

# Create directory if missing
mkdir -p "$TARGET_DIR"

# Copy manifest
cp "manifest.xml" "$TARGET_DIR/"

echo "Manifest installed to:"
echo "  $TARGET_DIR"
echo
echo "Restart Excel and check under:"
echo "  Insert → Add-ins → Shared Folder Add-ins"
