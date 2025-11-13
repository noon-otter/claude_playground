#!/bin/bash

TARGET_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

if [ -z "$1" ]; then
  echo "Usage: read_xl <filename>"
  exit 1
fi

cat "$TARGET_DIR/$1"
