using System;

namespace ExcelMerge
{
    /// <summary>
    /// Shared cell comparison logic used by both RowComparer (row-level LCS)
    /// and ExcelSheet.DiffCellsCaseEqual (cell-level diff).
    /// </summary>
    public static class CellComparer
    {
        public static string GetCompareValue(ExcelCell cell, bool compareFormula)
        {
            if (compareFormula && !string.IsNullOrEmpty(cell.Formula))
                return cell.Formula;

            return cell.Value;
        }

        public static bool AreEqual(ExcelCell src, ExcelCell dst, bool compareFormula, bool ignoreWhitespace, double numericPrecision)
        {
            if (src.Comment != dst.Comment)
                return false;

            var srcVal = GetCompareValue(src, compareFormula);
            var dstVal = GetCompareValue(dst, compareFormula);

            if (ignoreWhitespace)
            {
                srcVal = srcVal.Trim();
                dstVal = dstVal.Trim();
            }

            if (srcVal.Equals(dstVal))
                return true;

            if (numericPrecision > 0
                && double.TryParse(srcVal, out var srcNum)
                && double.TryParse(dstVal, out var dstNum))
            {
                return Math.Abs(srcNum - dstNum) <= numericPrecision;
            }

            return false;
        }

        public static string GetNormalizedHashValue(ExcelCell cell, bool compareFormula, bool ignoreWhitespace, double numericPrecision)
        {
            var value = GetCompareValue(cell, compareFormula);

            if (ignoreWhitespace)
                value = value.Trim();

            if (numericPrecision > 0 && double.TryParse(value, out var num))
            {
                var rounded = Math.Round(num / numericPrecision) * numericPrecision;
                return rounded.ToString("R");
            }

            return value;
        }
    }
}
