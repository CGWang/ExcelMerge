using System.Collections.Generic;
using System.Linq;
using System.Text;
using FastWpfGrid;
using ExcelMerge.GUI.Models;

namespace ExcelMerge.GUI.Services
{
    public static class LogBuilder
    {
        /// <summary>
        /// Builds a cell-based diff log from the selected cells, using configured log formats.
        /// </summary>
        /// <param name="srcModel">The source-side grid model.</param>
        /// <param name="dstModel">The destination-side grid model.</param>
        /// <param name="selectedCells">The cells currently selected in the grid.</param>
        /// <param name="modifiedLogFormat">Format string for modified cells (supports ${ROW}, ${COL}, ${LEFT}, ${RIGHT}).</param>
        /// <param name="addedLogFormat">Format string for added rows (supports ${ROW}).</param>
        /// <param name="removedLogFormat">Format string for removed rows (supports ${ROW}).</param>
        /// <param name="blankLabel">The label to use when a value is empty (e.g., "(blank)").</param>
        /// <returns>The formatted log string.</returns>
        public static string BuildCellBaseLog(
            DiffGridModel srcModel,
            DiffGridModel dstModel,
            IEnumerable<FastGridCellAddress> selectedCells,
            string modifiedLogFormat,
            string addedLogFormat,
            string removedLogFormat,
            string blankLabel)
        {
            if (srcModel == null || dstModel == null)
                return string.Empty;

            var builder = new StringBuilder();

            foreach (var row in selectedCells.GroupBy(c => c.Row))
            {
                var rowHeaderText = srcModel.GetRowHeaderText(row.Key.Value);
                if (string.IsNullOrEmpty(rowHeaderText))
                    rowHeaderText = dstModel.GetRowHeaderText(row.Key.Value);

                if (dstModel.IsAddedRow(row.Key.Value, true))
                {
                    var log = addedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText));

                    builder.AppendLine(log);
                    continue;
                }

                if (dstModel.IsRemovedRow(row.Key.Value, true))
                {
                    var log = removedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText));

                    builder.AppendLine(log);
                    continue;
                }

                foreach (var cell in row)
                {
                    if (cell.Row.Value == srcModel.ColumnHeaderIndex)
                        continue;

                    var srcText = srcModel.GetCellText(cell, true);
                    var dstText = dstModel.GetCellText(cell, true);
                    if (srcText == dstText)
                        continue;

                    var colHeaderText = srcModel.GetColumnHeaderText(cell.Column.Value);

                    if (string.IsNullOrEmpty(colHeaderText))
                        colHeaderText = dstModel.GetColumnHeaderText(cell.Column.Value);

                    if (string.IsNullOrEmpty(srcText))
                        srcText = blankLabel;

                    if (string.IsNullOrEmpty(dstText))
                        dstText = blankLabel;

                    if (string.IsNullOrEmpty(rowHeaderText))
                        rowHeaderText = blankLabel;

                    if (string.IsNullOrEmpty(colHeaderText))
                        colHeaderText = blankLabel;

                    var log = modifiedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText))
                        .Replace("${COL}", RemoveMultiLine(colHeaderText))
                        .Replace("${LEFT}", RemoveMultiLine(srcText))
                        .Replace("${RIGHT}", RemoveMultiLine(dstText));

                    builder.AppendLine(log);
                }
            }

            return builder.ToString();
        }

        /// <summary>
        /// Replaces newline sequences with spaces.
        /// </summary>
        public static string RemoveMultiLine(string text)
        {
            if (text == null)
                return string.Empty;

            return text.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        }
    }
}
