using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public class ExcelRow : IEquatable<ExcelRow>
    {
        public int Index { get; private set; }
        public List<ExcelCell> Cells { get; private set; }

        public ExcelRow(int index, IEnumerable<ExcelCell> cells)
        {
            Index = index;
            Cells = cells.ToList();
        }

        public override bool Equals(object obj)
        {
            var other = obj as ExcelRow;

            return Equals(other);
        }

        public override int GetHashCode()
        {
            var hash = 7;
            foreach (var cell in Cells)
            {
                hash = hash * 13 + cell.Value.GetHashCode();
            }

            return hash;
        }

        public bool Equals(ExcelRow other)
        {
            if (other == null)
                return false;

            if (Cells.Count != other.Cells.Count)
                return false;

            for (int i = 0; i < Cells.Count; i++)
            {
                if (Cells[i].Value != other.Cells[i].Value)
                    return false;
            }

            return true;
        }

        public bool IsBlank()
        {
            return Cells.All(c => string.IsNullOrEmpty(c.Value));
        }

        public void UpdateCells(IEnumerable<ExcelCell> cells)
        {
            Cells = cells.ToList();
        }
    }

    internal class RowComparer : IEqualityComparer<ExcelRow>
    {
        public HashSet<int> IgnoreColumns { get; private set; }
        public bool CompareFormula { get; private set; }
        public bool IgnoreWhitespace { get; private set; }
        public double NumericPrecision { get; private set; }

        public RowComparer(HashSet<int> ignoreColumns, bool compareFormula = false,
                           bool ignoreWhitespace = false, double numericPrecision = 0)
        {
            IgnoreColumns = ignoreColumns;
            CompareFormula = compareFormula;
            IgnoreWhitespace = ignoreWhitespace;
            NumericPrecision = numericPrecision;
        }

        public bool Equals(ExcelRow x, ExcelRow y)
        {
            if (x == null && y == null) return true;
            if (x == null || y == null) return false;
            if (x.Cells.Count != y.Cells.Count) return false;

            for (int i = 0; i < x.Cells.Count; i++)
            {
                if (IgnoreColumns.Contains(i))
                    continue;

                if (!AreCellsEqual(x.Cells[i], y.Cells[i]))
                    return false;
            }

            return true;
        }

        private bool AreCellsEqual(ExcelCell a, ExcelCell b)
        {
            var aVal = GetCellCompareValue(a);
            var bVal = GetCellCompareValue(b);

            if (IgnoreWhitespace)
            {
                aVal = aVal.Trim();
                bVal = bVal.Trim();
            }

            if (aVal == bVal)
                return true;

            if (NumericPrecision > 0
                && double.TryParse(aVal, out var aNum)
                && double.TryParse(bVal, out var bNum))
            {
                return Math.Abs(aNum - bNum) <= NumericPrecision;
            }

            return false;
        }

        public int GetHashCode(ExcelRow obj)
        {
            var hash = 7;
            for (int i = 0; i < obj.Cells.Count; i++)
            {
                if (IgnoreColumns.Contains(i))
                    continue;

                hash = hash * 13 + GetNormalizedHashValue(obj.Cells[i]).GetHashCode();
            }

            return hash;
        }

        private string GetNormalizedHashValue(ExcelCell cell)
        {
            var value = GetCellCompareValue(cell);

            if (IgnoreWhitespace)
                value = value.Trim();

            if (NumericPrecision > 0 && double.TryParse(value, out var num))
            {
                var rounded = Math.Round(num / NumericPrecision) * NumericPrecision;
                return rounded.ToString("R");
            }

            return value;
        }

        private string GetCellCompareValue(ExcelCell cell)
        {
            if (CompareFormula && !string.IsNullOrEmpty(cell.Formula))
                return cell.Formula;

            return cell.Value;
        }
    }
}
