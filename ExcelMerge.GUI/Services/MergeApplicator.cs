using System.Collections.Generic;
using System.Linq;
using FastWpfGrid;
using ExcelMerge.GUI.Models;

namespace ExcelMerge.GUI.Services
{
    /// <summary>
    /// Applies merge accept/reject decisions to a MergeResult based on the user's
    /// cell or row selection in the grid.  This class is UI-agnostic — it operates
    /// on a DiffGridModel and a set of selected cell addresses.
    /// </summary>
    public static class MergeApplicator
    {
        /// <summary>
        /// Accepts the chosen side for each individually selected cell.
        /// Uses TryGetCellDiffPublic to map visual coordinates to real diff coordinates.
        /// </summary>
        public static void ApplyToSelectedCells(
            MergeResult mergeResult,
            DiffGridModel model,
            IEnumerable<FastGridCellAddress> selectedCells,
            MergeSide side)
        {
            if (mergeResult == null || model == null || selectedCells == null)
                return;

            foreach (var cell in selectedCells)
            {
                if (!cell.IsCell) continue;

                ExcelCellDiff cellDiff;
                if (model.TryGetCellDiffPublic(cell.Row.Value, cell.Column.Value, out cellDiff))
                {
                    if (side == MergeSide.Src)
                        mergeResult.AcceptSrc(cellDiff.RowIndex, cellDiff.ColumnIndex);
                    else
                        mergeResult.AcceptDst(cellDiff.RowIndex, cellDiff.ColumnIndex);
                }
            }
        }

        /// <summary>
        /// Accepts the chosen side for every cell in each distinct selected row.
        /// Uses GetRealRowIndex to map visual rows to real diff rows.
        /// </summary>
        public static void ApplyToSelectedRows(
            MergeResult mergeResult,
            DiffGridModel model,
            IEnumerable<FastGridCellAddress> selectedCells,
            MergeSide side)
        {
            if (mergeResult == null || model == null || selectedCells == null)
                return;

            var rows = selectedCells
                .Where(c => c.IsCell)
                .Select(c => model.GetRealRowIndex(c.Row.Value))
                .Distinct();

            foreach (var row in rows)
            {
                if (side == MergeSide.Src)
                    mergeResult.AcceptSrcRow(row);
                else
                    mergeResult.AcceptDstRow(row);
            }
        }
    }
}
