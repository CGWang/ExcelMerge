# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ExcelMerge is a GUI diff tool for Excel files (.xls, .xlsx, .csv, .tsv), built with C# / WPF on .NET Framework 4.5.2. Forked from skanmera/ExcelMerge (MIT License). The goal is secondary development to add SVN (TortoiseSVN) integration, merge capabilities, and improved diff accuracy for a ~150-person team.

The development plan is documented in `.claude/EXCEL_MERGE_DEV_PLAN.md` (in Chinese).

## Build Commands

```bash
# Build entire solution (requires MSBuild / Visual Studio)
msbuild ExcelMerge.sln /p:Configuration=Debug /p:Platform="Any CPU"

# Build release
msbuild ExcelMerge.sln /p:Configuration=Release /p:Platform="Any CPU"

# Restore NuGet packages (if needed before build)
nuget restore ExcelMerge.sln

# Run diff algorithm tests (NetDiff.Test project)
# Uses MSTest — run via vstest.console or Visual Studio Test Explorer
vstest.console NetDiff/NetDiff.Test/bin/Debug/NetDiff.Test.dll
```

No custom build scripts exist. The installer project (ExcelMerge.Installer.vdproj) requires the legacy Visual Studio Setup Project extension.

## Architecture

### Solution Projects

| Project | Type | Purpose |
|---------|------|---------|
| **ExcelMerge** | Class Library | Core diff engine — file reading, sheet/row/cell models, diff logic |
| **ExcelMerge.GUI** | WPF Application | Main app — MVVM with Prism, CLI parsing, diff visualization |
| **FastWpfGrid** | Class Library | Custom high-performance virtualized grid control (core asset, modify with care) |
| **WriteableBitmapEx** | Class Library | Bitmap rendering support for FastWpfGrid |
| **NetDiff** | Class Library | Generic LCS-based diff algorithm |
| **NetDiff.Test** | Test Project | Unit tests for the diff algorithm |
| **ExcelMerge.ShellExtension** | COM DLL | Windows Explorer right-click context menu integration |
| **ExcelMerge.Installer** | Setup Project | MSI installer (WiX/vdproj) |

### Diff Flow

1. CLI parsing (`CommandLineOption.cs`) -> `CommandFactory` -> `DiffCommand`
2. `ExcelWorkbook.Load()` reads both files via NPOI
3. `ExcelSheet.Diff()` runs LCS-based diff via NetDiff's `DiffUtil`
4. Results stored in `ExcelSheetDiff` (contains `SortedDictionary<int, ExcelRowDiff>`)
5. `DiffViewModel` binds results to two synchronized `FastWpfGrid` controls in `DiffView.xaml`

### Key Dependencies

- **NPOI 2.3.0** — Excel file reading (does NOT require Microsoft Excel installed)
- **Prism.Wpf 6.3.0 + Unity 4.0.1** — MVVM framework and DI container
- **CommandLineParser 1.9.71** — CLI argument parsing
- **YamlDotNet 4.2.1** — Settings persistence
- **Extended.Wpf.Toolkit 3.2.0** — AvalonDock and WPF controls

### Data Model

```
ExcelWorkbook -> ExcelSheet -> SortedDictionary<int, ExcelRow> -> List<ExcelCell>
```

Cell status enum: `None | Modified | Added | Removed` (same pattern for row/column status).

`RowComparer` (`ExcelRow.cs`) implements `IEqualityComparer<ExcelRow>` with support for ignoring specific columns.

### Critical Components

- **FastWpfGrid**: Virtualized grid rendering only visible cells. Handles diff color rendering and synchronized left-right pane scrolling. Breaking the virtualization logic will cause severe performance issues with large files.
- **ExcelSheet.Diff()** (`ExcelSheet.cs`): Core diff logic — the main entry point for computing sheet-level diffs.
- **DiffView.xaml.cs** (~1100 lines): Main UI logic for the diff pane, event handling, grid model binding.

## CLI Usage

```bash
ExcelMerge.GUI diff -s <source_file> -d <dest_file> [-c <external_cmd>] [-i] [-w] [-v] [-e <empty_name>] [-k]
```

## Backward Compatibility

- Existing CLI arguments (`-s`, `-d`, `-c`, `-i`, `-w`, `-v`, `-e`, `-k`) must remain unchanged.
- New arguments should use `--long-option` style.
- Windows Explorer right-click integration must continue working.

## Known Issues

- Column insertion/deletion may display at incorrect positions (documented in README).
- Row alignment can be inaccurate with complex insertions/deletions (LCS algorithm limitation).
- No merge functionality (display-only diff).
- Reads cell values only — no formula, comment, or VBA comparison.
