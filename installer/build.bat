@echo off
echo === Publishing ExcelMerge ===
dotnet publish ..\ExcelMerge.GUI\ExcelMerge.GUI.csproj -c Release -r win-x64 --self-contained -p:PublishSingleFile=true

echo.
echo === Build installer with Inno Setup ===
echo Run: "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" ExcelMerge.iss
echo.
echo If Inno Setup is installed, the installer will be in installer\output\
pause
