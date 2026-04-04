using System;
using System.Linq;
using System.Windows;
using FastWpfGrid;
using ExcelMerge.GUI.Models;

namespace ExcelMerge.GUI.Services
{
    public static class ClipboardService
    {
        /// <summary>
        /// Copies the selected cells from the given grid to the clipboard,
        /// joining columns with the specified separator and rows with newlines.
        /// </summary>
        public static void CopySelectedCells(FastGridControl grid, string separator)
        {
            if (grid == null)
                return;

            var model = grid.Model as DiffGridModel;
            if (model == null)
                return;

            var cells = grid.SelectedCells;
            if (cells == null || !cells.Any())
                return;

            var tsv = string.Join(Environment.NewLine,
                cells
                    .GroupBy(c => c.Row.Value)
                    .OrderBy(g => g.Key)
                    .Select(g => string.Join(separator, g.Select(c => model.GetCellText(c, true)))));

            Clipboard.SetDataObject(tsv);
        }
    }
}
