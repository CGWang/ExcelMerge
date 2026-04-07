# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ExcelMerge is a GUI diff & merge tool for Excel files (.xls, .xlsx, .csv, .tsv), built with C# / WPF / .NET 8. Forked from skanmera/ExcelMerge (MIT License), with extensive secondary development for a ~150-person Unity game team using SVN (TortoiseSVN).

The development plan is documented in `.claude/EXCEL_MERGE_DEV_PLAN.md` (in Chinese).

## Build Commands

```bash
# Build entire solution
dotnet build ExcelMerge.sln

# Build release
dotnet build ExcelMerge.sln -c Release

# Run all tests (99 total: 31 NetDiff + 68 E2E)
dotnet test NetDiff/NetDiff.Test/NetDiff.Test.csproj
dotnet test E2ETest/E2ETest.csproj

# Publish self-contained single-file exe
dotnet publish ExcelMerge.GUI/ExcelMerge.GUI.csproj -c Release -r win-x64 --self-contained -p:PublishSingleFile=true

# Build installer (requires Inno Setup 6)
cd installer && build.bat
```

## CLI Usage

### Diff mode (TortoiseSVN diff)
```bash
ExcelMerge.GUI diff -s <source> -d <dest> [options]
ExcelMerge.GUI diff <source> <dest>                    # positional args

# TortoiseSVN config:
"path\ExcelMerge.GUI.exe" diff -s %base -d %mine --readonly-left --quit-on-close
```

### Merge mode (TortoiseSVN 3-way merge)
```bash
ExcelMerge.GUI merge --base <base> --mine <mine> --theirs <theirs> --output <merged>

# TortoiseSVN config:
"path\ExcelMerge.GUI.exe" merge --base %base --mine %mine --theirs %theirs --output %merged
```

## Architecture

### Solution Projects

| Project | Type | Purpose |
|---------|------|---------|
| **ExcelMerge** | Class Library | Core diff engine, cell/row/sheet models, comparison logic |
| **ExcelMerge.GUI** | WPF Application | Main app — CLI parsing, diff/merge visualization, MVVM |
| **FastWpfGrid** | Class Library | High-performance virtualized grid control (core asset) |
| **WriteableBitmapEx** | Class Library | Bitmap rendering support for FastWpfGrid |
| **NetDiff** | Class Library | Generic LCS-based diff algorithm |
| **NetDiff.Test** | Test Project | 31 unit tests for diff algorithm |
| **E2ETest** | Test Project | 68 E2E tests covering core engine, services, integration |
| **ExcelMerge.ShellExtension** | COM DLL | Windows Explorer right-click context menu |

### Core Library Key Files

| File | Purpose |
|------|---------|
| `ExcelSheet.cs` | 2-way LCS diff with column alignment |
| `ThreeWayDiff.cs` | 3-way merge engine (reuses 2-way alignment) |
| `CellComparer.cs` | Unified comparison logic (formula/whitespace/precision/comment) |
| `TextDiffUtil.cs` | Character-level inline diff for detail panel |
| `MergeWriter.cs` | Writes merged result to xlsx via NPOI |
| `MergeResult.cs` | 2-way merge decision tracking |

### GUI Services (extracted from DiffView)

| File | Purpose |
|------|---------|
| `Services/SearchService.cs` | Search history management |
| `Services/ClipboardService.cs` | Copy selected cells as TSV/CSV |
| `Services/LogBuilder.cs` | Diff log generation |
| `Services/DiffNavigator.cs` | Next/prev cell/row navigation |
| `Services/MergeApplicator.cs` | Accept src/dst merge operations |

### Data Model

```
ExcelWorkbook -> ExcelSheet -> SortedDictionary<int, ExcelRow> -> List<ExcelCell>

ExcelCell properties: Value, Formula, Comment
Comparison: CellComparer.AreEqual() (handles formula, whitespace, numeric precision, comment)

2-way diff result: ExcelSheetDiff -> ExcelRowDiff -> ExcelCellDiff (Status: None/Modified/Added/Removed)
3-way diff result: ThreeWayDiffResult -> CellMergeResult (Status: Unchanged/MineOnly/TheirsOnly/BothSame/Conflict)
```

### Key Dependencies

- **NPOI 2.7.2** — Excel file reading (no Microsoft Excel required)
- **Prism.Core 9.0.537** — MVVM base classes
- **CommandLineParser 2.9.1** — CLI argument parsing (multi-verb: diff, merge)
- **YamlDotNet 16.3.0** — Settings persistence
- **Extended.Wpf.Toolkit 4.6.1** — IntegerUpDown and ColorPicker controls

### 3-Panel Merge Layout

```
Default (2-panel diff):
┌──────────────┬──────────────┐
│  SRC (left)  │  DST (right) │
└──────────────┴──────────────┘

Merge mode with conflicts (3-panel, auto-expanded):
┌──────────┬──────────┬──────────┐
│ THEIRS   │  BASE    │  MINE    │
│ (left)   │ (center) │ (right)  │
└──────────┴──────────┴──────────┘
BASE panel: columns 3-4 in Grid, Width="0" by default, toggled via ShowBasePanel()
```

## Critical Components — Modify with Care

- **FastWpfGrid**: Virtualized rendering. Breaking it causes severe perf issues with large files.
- **DiffViewEventHandler**: Event dispatch system syncs scrolling/sizing across all grids. All `ResolveAll<FastGridControl>()` loops MUST guard with `grid.Model == null` check — BASE panel may have no model.
- **SimpleContainer.Resolve**: Returns `default` (not throw) when key not found. This is intentional — BASE panel doesn't register Rectangle/Grid for minimap.
- **ExcelSheet.Diff()**: Mutates input sheets (column shifting). ThreeWayDiff deep-copies sheets before each diff call.

## Development Lessons (Retrospective)

### Recurring Bug Pattern: BASE Panel Registration
The BASE panel was added to the existing 2-panel event system (container + dispatchers). Multiple crashes occurred because:
1. BASE was registered as a `FastGridControl` but lacked `Rectangle` and `Grid` (no minimap)
2. Event handlers iterated all registered grids without null-checking `Model`

**Rule**: When adding a new grid/panel to the event system, register ALL types that handlers expect to Resolve, or ensure ALL Resolve calls handle null gracefully.

### LCS Folding Creates "False Modified" Rows
`DiffUtil.OptimizeCaseDeletedFirst` folds adjacent Delete+Insert into Modified. This means:
- A row insertion can produce 1-2 "modified" rows at the boundary instead of pure Added
- 3-way merge with row insertion can create false conflicts at boundaries
- Tests should use tolerant assertions (`<= N` instead of `== 0`) for modified row counts

### ExcelSheet.Diff() Mutates Inputs
The column-shift logic in `ExcelSheet.Diff()` modifies `ExcelRow.Cells` in-place via `UpdateCells()`. This is dangerous when:
- The same sheet is used in multiple diff calls (ThreeWayDiff)
- Re-diffing after user changes config

**Rule**: Always deep-copy ExcelSheet before passing to Diff() if the sheet will be reused.

### Settings That Don't Work Are Bugs
`ColorModifiedRow` was disabled during the BeyondCompare-style highlighting refactor but the setting remained in the UI. Users could toggle it with no effect. Settings must either work or be removed from the UI.

### Constructor Signature Changes Break Silently
When ExcelCell went from `(value, formula, comment, colIdx, rowIdx)` to `(value, colIdx, rowIdx, formula, comment)`, existing call sites with positional args compiled fine but passed wrong values. Named or optional parameters are safer for constructors with many string args.

## Known Limitations

- LCS row alignment may produce boundary artifacts with complex insertions/deletions
- ThreeWayDiff row-level alignment depends on 2-way diff quality; row insertions may cause false conflicts at boundaries
- No formula evaluation — compares formula strings, not computed results
- No VBA, cell formatting, or named range comparison
- Column insertion/deletion display may be inaccurate (original upstream issue)

## Backward Compatibility

- Existing CLI arguments (`-s`, `-d`, `-c`, `-i`, `-w`, `-v`, `-e`, `-k`) unchanged
- New arguments use `--long-option` style
- Windows Explorer right-click integration works
- `merge` is a new verb, does not affect existing `diff` behavior
