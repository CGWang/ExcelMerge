using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelMerge
{
    public static class MergeWriter
    {
        /// <summary>
        /// Writes the merged result to an Excel file.
        /// Uses the BASE file as template (preserving formatting), then overwrites cell values
        /// with resolved values from the ThreeWayDiffResult.
        /// </summary>
        public static void Write(string baseFilePath, string outputPath, ThreeWayDiffResult mergeResult, string sheetName)
        {
            IWorkbook workbook;
            using (var fs = new FileStream(baseFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                workbook = WorkbookFactory.Create(fs);
            }

            var sheet = workbook.GetSheet(sheetName) ?? workbook.GetSheetAt(0);

            foreach (var rowEntry in mergeResult.Rows)
            {
                var rowIdx = rowEntry.Key;
                var row = sheet.GetRow(rowIdx) ?? sheet.CreateRow(rowIdx);

                foreach (var colEntry in rowEntry.Value)
                {
                    var colIdx = colEntry.Key;
                    var cellResult = colEntry.Value;

                    if (cellResult.Status == CellMergeStatus.Unchanged)
                        continue;

                    var resolvedValue = cellResult.ResolvedValue;
                    if (resolvedValue == null)
                        continue; // Unresolved conflict — skip

                    var cell = row.GetCell(colIdx) ?? row.CreateCell(colIdx);

                    // Try to set as numeric if possible
                    if (double.TryParse(resolvedValue, out var numValue))
                        cell.SetCellValue(numValue);
                    else
                        cell.SetCellValue(resolvedValue);
                }
            }

            using (var fs = new FileStream(outputPath, FileMode.Create))
            {
                workbook.Write(fs);
            }
        }
    }
}
