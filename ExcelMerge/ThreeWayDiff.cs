using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public enum CellMergeStatus
    {
        Unchanged,      // All three have the same value
        MineOnly,       // Only MINE differs from BASE
        TheirsOnly,     // Only THEIRS differs from BASE
        BothSame,       // Both MINE and THEIRS differ from BASE, but to the same value
        Conflict,       // Both differ from BASE, to different values
    }

    public class CellMergeResult
    {
        public int Row { get; }
        public int Column { get; }
        public string BaseValue { get; }
        public string MineValue { get; }
        public string TheirsValue { get; }
        public CellMergeStatus Status { get; }
        public string ResolvedValue { get; set; }  // mutable — user picks this for conflicts

        public CellMergeResult(int row, int column, string baseValue, string mineValue, string theirsValue)
        {
            Row = row;
            Column = column;
            BaseValue = baseValue ?? string.Empty;
            MineValue = mineValue ?? string.Empty;
            TheirsValue = theirsValue ?? string.Empty;

            // Determine status
            bool mineChanged = BaseValue != MineValue;
            bool theirsChanged = BaseValue != TheirsValue;

            if (!mineChanged && !theirsChanged)
            {
                Status = CellMergeStatus.Unchanged;
                ResolvedValue = BaseValue;
            }
            else if (mineChanged && !theirsChanged)
            {
                Status = CellMergeStatus.MineOnly;
                ResolvedValue = MineValue;
            }
            else if (!mineChanged && theirsChanged)
            {
                Status = CellMergeStatus.TheirsOnly;
                ResolvedValue = TheirsValue;
            }
            else if (MineValue == TheirsValue)
            {
                Status = CellMergeStatus.BothSame;
                ResolvedValue = MineValue;
            }
            else
            {
                Status = CellMergeStatus.Conflict;
                ResolvedValue = null; // Must be resolved by user
            }
        }
    }

    public class ThreeWayDiffResult
    {
        public SortedDictionary<int, SortedDictionary<int, CellMergeResult>> Rows { get; }
        public int ConflictCount { get; private set; }
        public int AutoMergedCount { get; private set; }
        public int TotalChangedCount { get; private set; }

        public ThreeWayDiffResult()
        {
            Rows = new SortedDictionary<int, SortedDictionary<int, CellMergeResult>>();
        }

        public void AddCell(CellMergeResult cell)
        {
            if (!Rows.ContainsKey(cell.Row))
                Rows[cell.Row] = new SortedDictionary<int, CellMergeResult>();

            Rows[cell.Row][cell.Column] = cell;

            if (cell.Status == CellMergeStatus.Conflict)
                ConflictCount++;
            if (cell.Status == CellMergeStatus.MineOnly || cell.Status == CellMergeStatus.TheirsOnly || cell.Status == CellMergeStatus.BothSame)
                AutoMergedCount++;
            if (cell.Status != CellMergeStatus.Unchanged)
                TotalChangedCount++;
        }

        public CellMergeResult GetCell(int row, int column)
        {
            if (Rows.TryGetValue(row, out var cols) && cols.TryGetValue(column, out var cell))
                return cell;
            return null;
        }

        public bool HasConflicts => ConflictCount > 0;

        public int UnresolvedConflictCount =>
            Rows.Values.SelectMany(r => r.Values)
                .Count(c => c.Status == CellMergeStatus.Conflict && c.ResolvedValue == null);

        public void ResolveConflict(int row, int column, string value)
        {
            var cell = GetCell(row, column);
            if (cell != null && cell.Status == CellMergeStatus.Conflict)
                cell.ResolvedValue = value;
        }
    }

    public static class ThreeWayDiff
    {
        /// <summary>
        /// Computes 3-way diff using LCS-based 2-way alignment.
        /// Runs ExcelSheet.Diff(base, mine) and ExcelSheet.Diff(base, theirs) to get
        /// proper row alignment, then walks the aligned results to determine per-cell
        /// 3-way merge status.
        ///
        /// Note: ExcelSheet.Diff() mutates its input sheets (column shifting for alignment).
        /// The base sheet is shared between both diffs, and the mine/theirs sheets may have
        /// already been mutated by a prior 2-way diff. To avoid corruption, all three sheets
        /// are deep-copied before each diff call.
        /// </summary>
        public static ThreeWayDiffResult Compute(ExcelSheet baseSheet, ExcelSheet mineSheet, ExcelSheet theirsSheet, ExcelSheetDiffConfig config = null)
        {
            var cfg = config ?? new ExcelSheetDiffConfig();

            // Deep-copy all sheets before each diff call since ExcelSheet.Diff() mutates inputs
            var baseCopyForMine = DeepCopySheet(baseSheet);
            var mineCopy = DeepCopySheet(mineSheet);
            var baseMineAligned = ExcelSheet.Diff(baseCopyForMine, mineCopy, cfg);

            var baseCopyForTheirs = DeepCopySheet(baseSheet);
            var theirsCopy = DeepCopySheet(theirsSheet);
            var baseTheirsAligned = ExcelSheet.Diff(baseCopyForTheirs, theirsCopy, cfg);

            // Walk aligned rows and compare cells
            var result = new ThreeWayDiffResult();

            var allRowIndices = new SortedSet<int>();
            foreach (var key in baseMineAligned.Rows.Keys) allRowIndices.Add(key);
            foreach (var key in baseTheirsAligned.Rows.Keys) allRowIndices.Add(key);

            foreach (var rowIdx in allRowIndices)
            {
                baseMineAligned.Rows.TryGetValue(rowIdx, out var mineRowDiff);
                baseTheirsAligned.Rows.TryGetValue(rowIdx, out var theirsRowDiff);

                var maxCols = Math.Max(
                    mineRowDiff?.Cells.Count ?? 0,
                    theirsRowDiff?.Cells.Count ?? 0);

                for (int col = 0; col < maxCols; col++)
                {
                    ExcelCellDiff mineCellDiff = null;
                    ExcelCellDiff theirsCellDiff = null;
                    mineRowDiff?.Cells.TryGetValue(col, out mineCellDiff);
                    theirsRowDiff?.Cells.TryGetValue(col, out theirsCellDiff);

                    // Base value comes from SrcCell of either diff (both share the same base)
                    var baseVal = mineCellDiff?.SrcCell?.Value ?? theirsCellDiff?.SrcCell?.Value ?? string.Empty;
                    var mineVal = mineCellDiff?.DstCell?.Value ?? baseVal;
                    var theirsVal = theirsCellDiff?.DstCell?.Value ?? baseVal;

                    result.AddCell(new CellMergeResult(rowIdx, col, baseVal, mineVal, theirsVal));
                }
            }

            return result;
        }

        private static ExcelSheet DeepCopySheet(ExcelSheet sheet)
        {
            var copy = new ExcelSheet();
            foreach (var kvp in sheet.Rows)
            {
                var copiedCells = kvp.Value.Cells.Select(c =>
                    new ExcelCell(c.Value, c.OriginalColumnIndex, c.OriginalRowIndex, c.Formula, c.Comment));
                copy.Rows.Add(kvp.Key, new ExcelRow(kvp.Value.Index, copiedCells));
            }
            return copy;
        }
    }
}
