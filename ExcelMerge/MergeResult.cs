using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelMerge
{
    public enum MergeSide
    {
        Src,
        Dst,
    }

    public class MergeResult
    {
        private readonly ExcelSheetDiff _diff;
        private readonly Dictionary<(int row, int col), MergeSide> _decisions = new();

        public MergeResult(ExcelSheetDiff diff)
        {
            _diff = diff;
        }

        public void AcceptSrc(int row, int col) => _decisions[(row, col)] = MergeSide.Src;
        public void AcceptDst(int row, int col) => _decisions[(row, col)] = MergeSide.Dst;

        public void AcceptSrcRow(int row)
        {
            if (!_diff.Rows.TryGetValue(row, out var rowDiff)) return;
            foreach (var col in rowDiff.Cells.Keys)
                _decisions[(row, col)] = MergeSide.Src;
        }

        public void AcceptDstRow(int row)
        {
            if (!_diff.Rows.TryGetValue(row, out var rowDiff)) return;
            foreach (var col in rowDiff.Cells.Keys)
                _decisions[(row, col)] = MergeSide.Dst;
        }

        public bool HasDecision(int row, int col) => _decisions.ContainsKey((row, col));

        public MergeSide? GetDecision(int row, int col)
        {
            return _decisions.TryGetValue((row, col), out var side) ? side : null;
        }

        public int DecisionCount => _decisions.Count;

        public string GetValue(int row, int col)
        {
            if (!_diff.Rows.TryGetValue(row, out var rowDiff)) return string.Empty;
            if (!rowDiff.Cells.TryGetValue(col, out var cellDiff)) return string.Empty;

            var side = _decisions.TryGetValue((row, col), out var s) ? s : MergeSide.Dst;
            return side == MergeSide.Src ? cellDiff.SrcCell.Value : cellDiff.DstCell.Value;
        }

        public void WriteToFile(string path)
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Merged");

            var maxRow = _diff.Rows.Any() ? _diff.Rows.Keys.Max() : -1;
            var maxCol = _diff.Rows.Values
                .Where(r => r.Cells.Any())
                .Select(r => r.Cells.Keys.Max())
                .DefaultIfEmpty(-1)
                .Max();

            for (int r = 0; r <= maxRow; r++)
            {
                var npoiRow = sheet.CreateRow(r);
                for (int c = 0; c <= maxCol; c++)
                {
                    var value = GetValue(r, c);
                    npoiRow.CreateCell(c).SetCellValue(value);
                }
            }

            using var fs = new FileStream(path, FileMode.Create);
            workbook.Write(fs);
        }
    }
}
