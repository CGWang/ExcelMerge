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
        /// Computes 3-way diff. All three sheets must have the same structure (same row/column count).
        /// For real usage, the sheets should be pre-aligned using the existing 2-way diff alignment.
        /// This simplified version compares cell values at matching positions.
        /// </summary>
        public static ThreeWayDiffResult Compute(ExcelSheet baseSheet, ExcelSheet mineSheet, ExcelSheet theirsSheet)
        {
            var result = new ThreeWayDiffResult();

            // Get all row indices from all three sheets
            var allRowIndices = new SortedSet<int>();
            foreach (var key in baseSheet.Rows.Keys) allRowIndices.Add(key);
            foreach (var key in mineSheet.Rows.Keys) allRowIndices.Add(key);
            foreach (var key in theirsSheet.Rows.Keys) allRowIndices.Add(key);

            foreach (var rowIdx in allRowIndices)
            {
                var baseRow = baseSheet.Rows.ContainsKey(rowIdx) ? baseSheet.Rows[rowIdx] : null;
                var mineRow = mineSheet.Rows.ContainsKey(rowIdx) ? mineSheet.Rows[rowIdx] : null;
                var theirsRow = theirsSheet.Rows.ContainsKey(rowIdx) ? theirsSheet.Rows[rowIdx] : null;

                var maxCols = Math.Max(
                    Math.Max(baseRow?.Cells.Count ?? 0, mineRow?.Cells.Count ?? 0),
                    theirsRow?.Cells.Count ?? 0);

                for (int col = 0; col < maxCols; col++)
                {
                    var baseVal = GetCellValue(baseRow, col);
                    var mineVal = GetCellValue(mineRow, col);
                    var theirsVal = GetCellValue(theirsRow, col);

                    var cellResult = new CellMergeResult(rowIdx, col, baseVal, mineVal, theirsVal);
                    result.AddCell(cellResult);
                }
            }

            return result;
        }

        private static string GetCellValue(ExcelRow row, int col)
        {
            if (row == null || col >= row.Cells.Count)
                return string.Empty;
            return row.Cells[col].Value;
        }
    }
}
