# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ExcelMerge is a GUI diff tool for Excel files (.xls, .xlsx, .csv, .tsv), built with C# / WPF. Forked from skanmera/ExcelMerge (MIT License), originally targeting .NET Framework 4.5.2 — upgrading to .NET 8 (LTS) as part of Phase 1 modernization. The goal is secondary development to add SVN (TortoiseSVN) integration, merge capabilities, and improved diff accuracy for a ~150-person Unity game team.

The development plan is documented in `.claude/EXCEL_MERGE_DEV_PLAN.md` (in Chinese).

## Build Commands

```bash
# Build entire solution
dotnet build ExcelMerge.sln

# Build release
dotnet build ExcelMerge.sln -c Release

# Run diff algorithm tests (NetDiff.Test — 31 MSTest tests)
dotnet test NetDiff/NetDiff.Test/NetDiff.Test.csproj

# Publish self-contained single-file exe
dotnet publish ExcelMerge.GUI/ExcelMerge.GUI.csproj -c Release --self-contained -p:PublishSingleFile=true
```

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

- **NPOI 2.7.2** — Excel file reading (does NOT require Microsoft Excel installed)
- **Prism.Core 9.0.537** — MVVM base classes (BindableBase, DelegateCommand)
- **CommandLineParser 2.9.1** — CLI argument parsing (verb-based API)
- **YamlDotNet 16.3.0** — Settings persistence
- **Extended.Wpf.Toolkit 4.6.1** — IntegerUpDown and ColorPicker controls
- **Microsoft.Xaml.Behaviors.Wpf 1.1.135** — WPF behaviors (replaces System.Windows.Interactivity)

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
