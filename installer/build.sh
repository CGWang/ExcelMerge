#!/bin/bash
set -e

PUBLISH_DIR="../ExcelMerge.GUI/bin/Release/net8.0-windows/win-x64/publish"

echo "========================================"
echo " ExcelMerge Release Build"
echo "========================================"

echo ""
echo "[1/2] Cleaning previous build..."
rm -rf "$PUBLISH_DIR" output/

echo ""
echo "[2/2] Publishing self-contained single-file..."
dotnet publish ../ExcelMerge.GUI/ExcelMerge.GUI.csproj -c Release -r win-x64 --self-contained -p:PublishSingleFile=true

echo ""
echo "========================================"
echo " Build complete!"
echo "========================================"
echo ""
echo "Published files: $PUBLISH_DIR"
echo ""
echo "To build installer, run on Windows:"
echo "  iscc ExcelMerge.iss"
