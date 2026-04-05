using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using ExcelMerge;

namespace E2ETest;

[TestClass]
public class CoreServiceTests
{
    private string _testDir;

    [TestInitialize]
    public void Setup()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"ExcelMerge_Core_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);
    }

    [TestCleanup]
    public void Cleanup()
    {
        try { Directory.Delete(_testDir, true); } catch { }
    }

    #region Helpers

    private string TestPath(string name) => Path.Combine(_testDir, name);

    private void CreateExcel(string path, Action<ISheet> populate, string sheetName = "Sheet1")
    {
        var wb = new XSSFWorkbook();
        var sheet = wb.CreateSheet(sheetName);
        populate(sheet);
        using var fs = new FileStream(path, FileMode.Create);
        wb.Write(fs);
    }

    private void CreateMultiSheetExcel(string path, params (string Name, Action<ISheet> Populate)[] sheets)
    {
        var wb = new XSSFWorkbook();
        foreach (var (name, populate) in sheets)
        {
            var sheet = wb.CreateSheet(name);
            populate(sheet);
        }
        using var fs = new FileStream(path, FileMode.Create);
        wb.Write(fs);
    }

    private void CreateXlsExcel(string path, Action<ISheet> populate)
    {
        var wb = new HSSFWorkbook();
        var sheet = wb.CreateSheet("Sheet1");
        populate(sheet);
        using var fs = new FileStream(path, FileMode.Create);
        wb.Write(fs);
    }

    private void SetCell(ISheet s, int row, int col, string value)
    {
        var r = s.GetRow(row) ?? s.CreateRow(row);
        (r.GetCell(col) ?? r.CreateCell(col)).SetCellValue(value);
    }

    private void SetNumeric(ISheet s, int row, int col, double value)
    {
        var r = s.GetRow(row) ?? s.CreateRow(row);
        (r.GetCell(col) ?? r.CreateCell(col)).SetCellValue(value);
    }

    private void SetFormula(ISheet s, int row, int col, string formula)
    {
        var r = s.GetRow(row) ?? s.CreateRow(row);
        (r.GetCell(col) ?? r.CreateCell(col)).SetCellFormula(formula);
    }

    private void CreateExcelWithComment(string path, int row, int col, string cellValue, string commentText)
    {
        var wb = new XSSFWorkbook();
        var sheet = wb.CreateSheet("Sheet1");
        var r = sheet.CreateRow(row);
        var cell = r.CreateCell(col);
        cell.SetCellValue(cellValue);

        var drawing = sheet.CreateDrawingPatriarch();
        var anchor = drawing.CreateAnchor(0, 0, 0, 0, col, row, col + 2, row + 1);
        var comment = drawing.CreateCellComment(anchor);
        comment.String = new XSSFRichTextString(commentText);
        cell.CellComment = comment;

        using var fs = new FileStream(path, FileMode.Create);
        wb.Write(fs);
    }

    private ThreeWayDiffResult ComputeThreeWay(string basePath, string minePath, string theirsPath)
    {
        var config = new ExcelSheetReadConfig();
        var baseWb = ExcelWorkbook.Create(basePath, config);
        var mineWb = ExcelWorkbook.Create(minePath, config);
        var theirsWb = ExcelWorkbook.Create(theirsPath, config);

        return ThreeWayDiff.Compute(
            baseWb.Sheets["Sheet1"], mineWb.Sheets["Sheet1"], theirsWb.Sheets["Sheet1"]);
    }

    #endregion

    // ═══════════════════════════════════════════════════════════
    // 1. CellComparer edge cases
    // ═══════════════════════════════════════════════════════════

    [TestMethod]
    public void CellComparer_FormulaMode_SameValueDifferentFormula()
    {
        // Two cells with the same display value but different formulas
        var a = new ExcelCell("20", 0, 0, formula: "=A1*2");
        var b = new ExcelCell("20", 0, 0, formula: "=A1+10");

        // Without formula comparison, they are equal (same display value)
        Assert.IsTrue(CellComparer.AreEqual(a, b, compareFormula: false, ignoreWhitespace: false, numericPrecision: 0),
            "Same value, formula comparison OFF -> equal");

        // With formula comparison, they differ
        Assert.IsFalse(CellComparer.AreEqual(a, b, compareFormula: true, ignoreWhitespace: false, numericPrecision: 0),
            "Same value, different formula, formula comparison ON -> not equal");
    }

    [TestMethod]
    public void CellComparer_FormulaMode_SameFormula()
    {
        var a = new ExcelCell("20", 0, 0, formula: "=A1*2");
        var b = new ExcelCell("20", 0, 0, formula: "=A1*2");

        Assert.IsTrue(CellComparer.AreEqual(a, b, compareFormula: true, ignoreWhitespace: false, numericPrecision: 0),
            "Same value and same formula -> equal");
    }

    [TestMethod]
    public void CellComparer_CommentComparison_SameValueSameFormulaDifferentComment()
    {
        var a = new ExcelCell("100", 0, 0, formula: "=SUM(A1:A10)", comment: "Total sales");
        var b = new ExcelCell("100", 0, 0, formula: "=SUM(A1:A10)", comment: "Grand total");

        // Comments differ -> not equal even when value and formula match
        Assert.IsFalse(CellComparer.AreEqual(a, b, compareFormula: true, ignoreWhitespace: false, numericPrecision: 0),
            "Different comments -> not equal regardless of value/formula match");
    }

    [TestMethod]
    public void CellComparer_CommentComparison_BothNoComment()
    {
        var a = new ExcelCell("hello", 0, 0);
        var b = new ExcelCell("hello", 0, 0);

        Assert.IsTrue(CellComparer.AreEqual(a, b, compareFormula: false, ignoreWhitespace: false, numericPrecision: 0),
            "No comments on either -> equal");
    }

    [TestMethod]
    public void CellComparer_CommentComparison_OneHasComment()
    {
        var a = new ExcelCell("hello", 0, 0, comment: "a note");
        var b = new ExcelCell("hello", 0, 0);

        Assert.IsFalse(CellComparer.AreEqual(a, b, compareFormula: false, ignoreWhitespace: false, numericPrecision: 0),
            "One cell has comment, other does not -> not equal");
    }

    [TestMethod]
    public void CellComparer_NumericPrecision_NegativeNumbers()
    {
        var a = new ExcelCell("-3.14", 0, 0);
        var b = new ExcelCell("-3.15", 0, 0);

        Assert.IsFalse(CellComparer.AreEqual(a, b, false, false, 0),
            "Exact: -3.14 != -3.15");
        Assert.IsTrue(CellComparer.AreEqual(a, b, false, false, 0.02),
            "Tolerance 0.02: -3.14 approx -3.15");
    }

    [TestMethod]
    public void CellComparer_NumericPrecision_LargeNegativeNumbers()
    {
        var a = new ExcelCell("-1000.001", 0, 0);
        var b = new ExcelCell("-1000.002", 0, 0);

        Assert.IsTrue(CellComparer.AreEqual(a, b, false, false, 0.01),
            "Tolerance 0.01: -1000.001 approx -1000.002");
        Assert.IsFalse(CellComparer.AreEqual(a, b, false, false, 0.0001),
            "Tolerance 0.0001: -1000.001 != -1000.002");
    }

    [TestMethod]
    public void CellComparer_NumericPrecision_NonNumericStringsFallBackToExact()
    {
        var a = new ExcelCell("abc", 0, 0);
        var b = new ExcelCell("abd", 0, 0);

        // Even with numeric precision set, non-numeric strings should not be "approximately equal"
        Assert.IsFalse(CellComparer.AreEqual(a, b, false, false, 100),
            "Non-numeric strings fall back to exact comparison, 'abc' != 'abd'");
    }

    [TestMethod]
    public void CellComparer_NumericPrecision_OneNumericOneNot()
    {
        var a = new ExcelCell("3.14", 0, 0);
        var b = new ExcelCell("pi", 0, 0);

        Assert.IsFalse(CellComparer.AreEqual(a, b, false, false, 100),
            "One numeric, one not -> not equal even with large tolerance");
    }

    [TestMethod]
    public void CellComparer_TrimMode_EmptyVsWhitespace()
    {
        var empty = new ExcelCell("", 0, 0);
        var spaces = new ExcelCell("   ", 0, 0);
        var tabs = new ExcelCell("\t\t", 0, 0);

        Assert.IsFalse(CellComparer.AreEqual(empty, spaces, false, false, 0),
            "Exact mode: empty != spaces");
        Assert.IsTrue(CellComparer.AreEqual(empty, spaces, false, true, 0),
            "Trim mode: empty == spaces (both trim to empty)");
        Assert.IsTrue(CellComparer.AreEqual(empty, tabs, false, true, 0),
            "Trim mode: empty == tabs (both trim to empty)");
        Assert.IsTrue(CellComparer.AreEqual(spaces, tabs, false, true, 0),
            "Trim mode: spaces == tabs (both trim to empty)");
    }

    [TestMethod]
    public void CellComparer_TrimMode_LeadingTrailingOnly()
    {
        var a = new ExcelCell("  hello world  ", 0, 0);
        var b = new ExcelCell("hello world", 0, 0);

        Assert.IsTrue(CellComparer.AreEqual(a, b, false, true, 0),
            "Trim mode: leading/trailing whitespace ignored");

        // Internal whitespace should still matter
        var c = new ExcelCell("hello  world", 0, 0);
        Assert.IsFalse(CellComparer.AreEqual(b, c, false, true, 0),
            "Trim mode: internal whitespace difference still matters");
    }

    [TestMethod]
    public void CellComparer_GetNormalizedHashValue_ConsistentWithAreEqual()
    {
        var a = new ExcelCell("3.14", 0, 0);
        var b = new ExcelCell("3.15", 0, 0);

        // With tolerance 0.1, 3.14 and 3.15 should be equal
        Assert.IsTrue(CellComparer.AreEqual(a, b, false, false, 0.1));

        // Their normalized hash values should also match
        var hashA = CellComparer.GetNormalizedHashValue(a, false, false, 0.1);
        var hashB = CellComparer.GetNormalizedHashValue(b, false, false, 0.1);
        Assert.AreEqual(hashA, hashB,
            "Normalized hash values should be the same for values within tolerance");
    }

    [TestMethod]
    public void CellComparer_GetCompareValue_FormulaVsValue()
    {
        var cellWithFormula = new ExcelCell("100", 0, 0, formula: "=A1*10");
        var cellNoFormula = new ExcelCell("100", 0, 0);

        Assert.AreEqual("100", CellComparer.GetCompareValue(cellNoFormula, compareFormula: true),
            "Cell without formula returns value even in formula mode");
        Assert.AreEqual("=A1*10", CellComparer.GetCompareValue(cellWithFormula, compareFormula: true),
            "Cell with formula returns formula in formula mode");
        Assert.AreEqual("100", CellComparer.GetCompareValue(cellWithFormula, compareFormula: false),
            "Cell with formula returns value when formula mode OFF");
    }

    // ═══════════════════════════════════════════════════════════
    // 2. TextDiffUtil edge cases
    // ═══════════════════════════════════════════════════════════

    [TestMethod]
    public void TextDiffUtil_LongStringWithSingleCharDiff()
    {
        var longStr = new string('A', 100);
        var modified = longStr.Substring(0, 50) + "B" + longStr.Substring(51);

        var srcSegs = TextDiffUtil.ComputeInlineDiffSrc(longStr, modified);
        var dstSegs = TextDiffUtil.ComputeInlineDiff(longStr, modified);

        // Reconstructed texts must match originals
        var srcText = string.Concat(srcSegs.Select(s => s.Text));
        var dstText = string.Concat(dstSegs.Select(s => s.Text));
        Assert.AreEqual(longStr, srcText, "Src reconstruction for long string");
        Assert.AreEqual(modified, dstText, "Dst reconstruction for long string");

        // Should have both modified and unmodified segments
        Assert.IsTrue(srcSegs.Any(s => s.IsModified), "Has modified segment in src");
        Assert.IsTrue(srcSegs.Any(s => !s.IsModified), "Has unmodified segment in src");

        // Total modified characters should be small (just the one char difference region)
        var modifiedCharCount = srcSegs.Where(s => s.IsModified).Sum(s => s.Text.Length);
        Assert.IsTrue(modifiedCharCount <= 5, $"Small modification region (got {modifiedCharCount} modified chars)");
    }

    [TestMethod]
    public void TextDiffUtil_CompletelyDifferentStrings()
    {
        var src = "AAAA";
        var dst = "ZZZZ";

        var srcSegs = TextDiffUtil.ComputeInlineDiffSrc(src, dst);
        var dstSegs = TextDiffUtil.ComputeInlineDiff(src, dst);

        var srcText = string.Concat(srcSegs.Select(s => s.Text));
        var dstText = string.Concat(dstSegs.Select(s => s.Text));
        Assert.AreEqual(src, srcText, "Src reconstruction");
        Assert.AreEqual(dst, dstText, "Dst reconstruction");

        // All segments should be modified since the strings are completely different
        Assert.IsTrue(srcSegs.All(s => s.IsModified), "All src segments are modified");
        Assert.IsTrue(dstSegs.All(s => s.IsModified), "All dst segments are modified");
    }

    [TestMethod]
    public void TextDiffUtil_UnicodeCjkCharacters()
    {
        var src = "Hello 世界 Test";
        var dst = "Hello 地球 Test";

        var srcSegs = TextDiffUtil.ComputeInlineDiffSrc(src, dst);
        var dstSegs = TextDiffUtil.ComputeInlineDiff(src, dst);

        var srcText = string.Concat(srcSegs.Select(s => s.Text));
        var dstText = string.Concat(dstSegs.Select(s => s.Text));
        Assert.AreEqual(src, srcText, "Src reconstruction with CJK");
        Assert.AreEqual(dst, dstText, "Dst reconstruction with CJK");

        // "Hello " and " Test" should be unmodified
        Assert.IsTrue(srcSegs.Any(s => !s.IsModified && s.Text.Contains("Hello")),
            "Common prefix 'Hello' is unmodified");
        Assert.IsTrue(dstSegs.Any(s => !s.IsModified && s.Text.Contains("Test")),
            "Common suffix 'Test' is unmodified");
    }

    [TestMethod]
    public void TextDiffUtil_PureChineseStrings()
    {
        var src = "你好世界";
        var dst = "你好地球";

        var srcSegs = TextDiffUtil.ComputeInlineDiffSrc(src, dst);
        var dstSegs = TextDiffUtil.ComputeInlineDiff(src, dst);

        var srcText = string.Concat(srcSegs.Select(s => s.Text));
        var dstText = string.Concat(dstSegs.Select(s => s.Text));
        Assert.AreEqual(src, srcText, "Src reconstruction with pure Chinese");
        Assert.AreEqual(dst, dstText, "Dst reconstruction with pure Chinese");

        // "你好" should be unmodified on both sides
        Assert.IsTrue(srcSegs.Any(s => !s.IsModified), "Has unmodified Chinese chars in src");
        Assert.IsTrue(dstSegs.Any(s => !s.IsModified), "Has unmodified Chinese chars in dst");
    }

    [TestMethod]
    public void TextDiffUtil_MultiLineText()
    {
        var src = "Line1\nLine2\nLine3";
        var dst = "Line1\nModified\nLine3";

        var srcSegs = TextDiffUtil.ComputeInlineDiffSrc(src, dst);
        var dstSegs = TextDiffUtil.ComputeInlineDiff(src, dst);

        var srcText = string.Concat(srcSegs.Select(s => s.Text));
        var dstText = string.Concat(dstSegs.Select(s => s.Text));
        Assert.AreEqual(src, srcText, "Src multiline reconstruction");
        Assert.AreEqual(dst, dstText, "Dst multiline reconstruction");

        // Line1 and Line3 should be in unmodified segments
        Assert.IsTrue(srcSegs.Any(s => !s.IsModified && s.Text.Contains("Line1")),
            "Line1 is unmodified");
        Assert.IsTrue(dstSegs.Any(s => !s.IsModified && s.Text.Contains("Line3")),
            "Line3 is unmodified");
    }

    [TestMethod]
    public void TextDiffUtil_NullInputs()
    {
        // null src
        var dstSegs = TextDiffUtil.ComputeInlineDiff(null, "abc");
        var dstText = string.Concat(dstSegs.Select(s => s.Text));
        Assert.AreEqual("abc", dstText, "null src -> dst segments contain 'abc'");
        Assert.IsTrue(dstSegs.All(s => s.IsModified), "All dst segments are modified (inserted)");

        // null dst
        var srcSegs = TextDiffUtil.ComputeInlineDiffSrc("abc", null);
        var srcText = string.Concat(srcSegs.Select(s => s.Text));
        Assert.AreEqual("abc", srcText, "null dst -> src segments contain 'abc'");
        Assert.IsTrue(srcSegs.All(s => s.IsModified), "All src segments are modified (deleted)");
    }

    [TestMethod]
    public void TextDiffUtil_SpecialCharacters()
    {
        var src = "price: $100.00 (USD)";
        var dst = "price: $200.00 (EUR)";

        var srcSegs = TextDiffUtil.ComputeInlineDiffSrc(src, dst);
        var dstSegs = TextDiffUtil.ComputeInlineDiff(src, dst);

        var srcText = string.Concat(srcSegs.Select(s => s.Text));
        var dstText = string.Concat(dstSegs.Select(s => s.Text));
        Assert.AreEqual(src, srcText, "Src with special chars");
        Assert.AreEqual(dst, dstText, "Dst with special chars");
    }

    // ═══════════════════════════════════════════════════════════
    // 3. ThreeWayDiff additional scenarios
    // ═══════════════════════════════════════════════════════════

    [TestMethod]
    public void ThreeWayDiff_BothSameChange_BothSameStatus()
    {
        var basePath = TestPath("both_base.xlsx");
        var minePath = TestPath("both_mine.xlsx");
        var theirsPath = TestPath("both_theirs.xlsx");

        CreateExcel(basePath, s =>
        {
            SetCell(s, 0, 0, "A"); SetCell(s, 0, 1, "original");
        });
        // Both sides make the same change
        CreateExcel(minePath, s =>
        {
            SetCell(s, 0, 0, "A"); SetCell(s, 0, 1, "updated");
        });
        CreateExcel(theirsPath, s =>
        {
            SetCell(s, 0, 0, "A"); SetCell(s, 0, 1, "updated");
        });

        var result = ComputeThreeWay(basePath, minePath, theirsPath);

        Assert.AreEqual(0, result.ConflictCount, "No conflicts when both make same change");

        // Find the cell that was changed
        var changedCells = result.Rows.Values
            .SelectMany(r => r.Values)
            .Where(c => c.Status == CellMergeStatus.BothSame)
            .ToList();

        Assert.IsTrue(changedCells.Count >= 1, $"At least 1 BothSame cell (got {changedCells.Count})");
        var cell = changedCells.First();
        Assert.AreEqual("updated", cell.ResolvedValue, "BothSame resolved to the shared value");
    }

    [TestMethod]
    public void ThreeWayDiff_AllUnchanged_UnchangedStatus()
    {
        var basePath = TestPath("unch_base.xlsx");
        var minePath = TestPath("unch_mine.xlsx");
        var theirsPath = TestPath("unch_theirs.xlsx");

        CreateExcel(basePath, s =>
        {
            SetCell(s, 0, 0, "A"); SetCell(s, 0, 1, "value");
            SetCell(s, 1, 0, "B"); SetCell(s, 1, 1, "data");
        });
        // Both sides identical to base
        File.Copy(basePath, minePath);
        File.Copy(basePath, theirsPath);

        var result = ComputeThreeWay(basePath, minePath, theirsPath);

        Assert.AreEqual(0, result.ConflictCount, "No conflicts");
        Assert.AreEqual(0, result.TotalChangedCount, "No changes at all");
        Assert.IsFalse(result.HasConflicts);

        // All cells should be Unchanged
        var allCells = result.Rows.Values.SelectMany(r => r.Values).ToList();
        Assert.IsTrue(allCells.All(c => c.Status == CellMergeStatus.Unchanged),
            "All cells are Unchanged");
    }

    [TestMethod]
    public void ThreeWayDiff_ResolveConflict_VerifyResolvedValue()
    {
        var basePath = TestPath("resolve_base.xlsx");
        var minePath = TestPath("resolve_mine.xlsx");
        var theirsPath = TestPath("resolve_theirs.xlsx");

        CreateExcel(basePath, s =>
        {
            SetCell(s, 0, 0, "key"); SetCell(s, 0, 1, "original");
        });
        CreateExcel(minePath, s =>
        {
            SetCell(s, 0, 0, "key"); SetCell(s, 0, 1, "mine_version");
        });
        CreateExcel(theirsPath, s =>
        {
            SetCell(s, 0, 0, "key"); SetCell(s, 0, 1, "theirs_version");
        });

        var result = ComputeThreeWay(basePath, minePath, theirsPath);

        Assert.IsTrue(result.HasConflicts, "Has conflicts");
        Assert.IsTrue(result.UnresolvedConflictCount >= 1, "Has unresolved conflicts");

        // Find the conflict cell and resolve it
        var conflictCell = result.Rows.Values
            .SelectMany(r => r.Values)
            .First(c => c.Status == CellMergeStatus.Conflict);

        Assert.IsNull(conflictCell.ResolvedValue, "Initially unresolved");

        // Resolve the conflict
        result.ResolveConflict(conflictCell.Row, conflictCell.Column, "manually_resolved");

        Assert.AreEqual("manually_resolved", conflictCell.ResolvedValue,
            "ResolvedValue set to manually chosen value");
        Assert.AreEqual(0, result.UnresolvedConflictCount,
            "No more unresolved conflicts after resolution");
    }

    [TestMethod]
    public void ThreeWayDiff_LargeFile_100PlusRows()
    {
        var basePath = TestPath("large3_base.xlsx");
        var minePath = TestPath("large3_mine.xlsx");
        var theirsPath = TestPath("large3_theirs.xlsx");

        CreateExcel(basePath, s =>
        {
            for (int i = 0; i < 120; i++)
            {
                SetCell(s, i, 0, $"Row{i}");
                SetCell(s, i, 1, $"Base{i}");
            }
        });
        // MINE changes rows 10, 20, 30
        CreateExcel(minePath, s =>
        {
            for (int i = 0; i < 120; i++)
            {
                SetCell(s, i, 0, $"Row{i}");
                if (i == 10 || i == 20 || i == 30)
                    SetCell(s, i, 1, $"Mine{i}");
                else
                    SetCell(s, i, 1, $"Base{i}");
            }
        });
        // THEIRS changes rows 50, 60, 70
        CreateExcel(theirsPath, s =>
        {
            for (int i = 0; i < 120; i++)
            {
                SetCell(s, i, 0, $"Row{i}");
                if (i == 50 || i == 60 || i == 70)
                    SetCell(s, i, 1, $"Theirs{i}");
                else
                    SetCell(s, i, 1, $"Base{i}");
            }
        });

        var result = ComputeThreeWay(basePath, minePath, theirsPath);

        // No conflicts since mine and theirs change different rows
        Assert.AreEqual(0, result.ConflictCount, "No conflicts when different rows changed");
        Assert.IsTrue(result.AutoMergedCount >= 6, $"At least 6 auto-merged cells (got {result.AutoMergedCount})");
        Assert.IsTrue(result.TotalChangedCount >= 6, $"At least 6 total changed (got {result.TotalChangedCount})");
    }

    [TestMethod]
    public void ThreeWayDiff_MineOnly_Status()
    {
        var basePath = TestPath("mineonly_base.xlsx");
        var minePath = TestPath("mineonly_mine.xlsx");
        var theirsPath = TestPath("mineonly_theirs.xlsx");

        CreateExcel(basePath, s =>
        {
            SetCell(s, 0, 0, "ID"); SetCell(s, 0, 1, "Value");
        });
        CreateExcel(minePath, s =>
        {
            SetCell(s, 0, 0, "ID"); SetCell(s, 0, 1, "Changed");
        });
        // Theirs same as base
        File.Copy(basePath, theirsPath);

        var result = ComputeThreeWay(basePath, minePath, theirsPath);

        Assert.AreEqual(0, result.ConflictCount);
        var mineOnlyCells = result.Rows.Values
            .SelectMany(r => r.Values)
            .Where(c => c.Status == CellMergeStatus.MineOnly)
            .ToList();
        Assert.IsTrue(mineOnlyCells.Count >= 1, "At least 1 MineOnly cell");
        Assert.AreEqual("Changed", mineOnlyCells.First().ResolvedValue,
            "MineOnly resolved to MINE value");
    }

    [TestMethod]
    public void ThreeWayDiff_TheirsOnly_Status()
    {
        var basePath = TestPath("theirsonly_base.xlsx");
        var minePath = TestPath("theirsonly_mine.xlsx");
        var theirsPath = TestPath("theirsonly_theirs.xlsx");

        CreateExcel(basePath, s =>
        {
            SetCell(s, 0, 0, "ID"); SetCell(s, 0, 1, "Value");
        });
        // Mine same as base
        File.Copy(basePath, minePath);
        CreateExcel(theirsPath, s =>
        {
            SetCell(s, 0, 0, "ID"); SetCell(s, 0, 1, "TheirChange");
        });

        var result = ComputeThreeWay(basePath, minePath, theirsPath);

        Assert.AreEqual(0, result.ConflictCount);
        var theirsOnlyCells = result.Rows.Values
            .SelectMany(r => r.Values)
            .Where(c => c.Status == CellMergeStatus.TheirsOnly)
            .ToList();
        Assert.IsTrue(theirsOnlyCells.Count >= 1, "At least 1 TheirsOnly cell");
        Assert.AreEqual("TheirChange", theirsOnlyCells.First().ResolvedValue,
            "TheirsOnly resolved to THEIRS value");
    }

    // ═══════════════════════════════════════════════════════════
    // 4. MergeWriter edge cases
    // ═══════════════════════════════════════════════════════════

    [TestMethod]
    public void MergeWriter_NoChanges_OutputMatchesBase()
    {
        var basePath = TestPath("mw_nochange_base.xlsx");
        var outputPath = TestPath("mw_nochange_out.xlsx");

        CreateExcel(basePath, s =>
        {
            SetCell(s, 0, 0, "ID"); SetCell(s, 0, 1, "Name"); SetCell(s, 0, 2, "Score");
            SetCell(s, 1, 0, "1");  SetCell(s, 1, 1, "Alice"); SetCell(s, 1, 2, "90");
            SetCell(s, 2, 0, "2");  SetCell(s, 2, 1, "Bob");   SetCell(s, 2, 2, "85");
        });

        // Empty merge result — no changes at all
        var mergeResult = new ThreeWayDiffResult();

        MergeWriter.Write(basePath, outputPath, mergeResult, "Sheet1");

        Assert.IsTrue(File.Exists(outputPath), "Output file created");

        // Verify output matches base
        var config = new ExcelSheetReadConfig();
        var baseWb = ExcelWorkbook.Create(basePath, config);
        var outWb = ExcelWorkbook.Create(outputPath, config);

        Assert.AreEqual(baseWb.Sheets["Sheet1"].Rows.Count, outWb.Sheets["Sheet1"].Rows.Count,
            "Same row count");

        foreach (var kvp in baseWb.Sheets["Sheet1"].Rows)
        {
            var baseRow = kvp.Value;
            var outRow = outWb.Sheets["Sheet1"].Rows[kvp.Key];
            for (int c = 0; c < baseRow.Cells.Count; c++)
            {
                Assert.AreEqual(baseRow.Cells[c].Value, outRow.Cells[c].Value,
                    $"Cell [{kvp.Key},{c}] matches base");
            }
        }
    }

    [TestMethod]
    public void MergeWriter_NumericValues_WrittenAsNumbers()
    {
        var basePath = TestPath("mw_numeric_base.xlsx");
        var outputPath = TestPath("mw_numeric_out.xlsx");

        CreateExcel(basePath, s =>
        {
            SetCell(s, 0, 0, "Value");
            SetNumeric(s, 1, 0, 100);
        });

        var mergeResult = new ThreeWayDiffResult();
        mergeResult.AddCell(new CellMergeResult(1, 0, "100", "200", "100"));

        MergeWriter.Write(basePath, outputPath, mergeResult, "Sheet1");

        // Read back with NPOI directly to check cell type
        IWorkbook wb;
        using (var fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            wb = WorkbookFactory.Create(fs);
        }
        var cell = wb.GetSheet("Sheet1").GetRow(1).GetCell(0);
        Assert.AreEqual(CellType.Numeric, cell.CellType, "Numeric value written as numeric cell type");
        Assert.AreEqual(200.0, cell.NumericCellValue, 0.001, "Numeric value is 200");
    }

    [TestMethod]
    public void MergeWriter_UnresolvedConflicts_Skipped()
    {
        var basePath = TestPath("mw_unresolved_base.xlsx");
        var outputPath = TestPath("mw_unresolved_out.xlsx");

        CreateExcel(basePath, s =>
        {
            SetCell(s, 0, 0, "ID"); SetCell(s, 0, 1, "Value");
            SetCell(s, 1, 0, "1");  SetCell(s, 1, 1, "original");
        });

        // Create a conflict with null ResolvedValue
        var mergeResult = new ThreeWayDiffResult();
        var conflictCell = new CellMergeResult(1, 1, "original", "mine_ver", "theirs_ver");
        Assert.AreEqual(CellMergeStatus.Conflict, conflictCell.Status);
        Assert.IsNull(conflictCell.ResolvedValue, "Conflict has null ResolvedValue");
        mergeResult.AddCell(conflictCell);

        MergeWriter.Write(basePath, outputPath, mergeResult, "Sheet1");

        // The original value should remain (unresolved conflict is skipped)
        var outWb = ExcelWorkbook.Create(outputPath, new ExcelSheetReadConfig());
        var outCell = outWb.Sheets["Sheet1"].Rows[1].Cells[1];
        Assert.AreEqual("original", outCell.Value,
            "Unresolved conflict leaves original value unchanged");
    }

    [TestMethod]
    public void MergeWriter_StringValue_WrittenAsString()
    {
        var basePath = TestPath("mw_string_base.xlsx");
        var outputPath = TestPath("mw_string_out.xlsx");

        CreateExcel(basePath, s =>
        {
            SetCell(s, 0, 0, "Name");
            SetCell(s, 1, 0, "Alice");
        });

        var mergeResult = new ThreeWayDiffResult();
        mergeResult.AddCell(new CellMergeResult(1, 0, "Alice", "Bob", "Alice"));

        MergeWriter.Write(basePath, outputPath, mergeResult, "Sheet1");

        var outWb = ExcelWorkbook.Create(outputPath, new ExcelSheetReadConfig());
        Assert.AreEqual("Bob", outWb.Sheets["Sheet1"].Rows[1].Cells[0].Value,
            "String value written correctly");
    }

    [TestMethod]
    public void MergeWriter_MultipleChanges_AllApplied()
    {
        var basePath = TestPath("mw_multi_base.xlsx");
        var outputPath = TestPath("mw_multi_out.xlsx");

        CreateExcel(basePath, s =>
        {
            SetCell(s, 0, 0, "A"); SetCell(s, 0, 1, "B"); SetCell(s, 0, 2, "C");
            SetCell(s, 1, 0, "1"); SetCell(s, 1, 1, "2"); SetCell(s, 1, 2, "3");
            SetCell(s, 2, 0, "4"); SetCell(s, 2, 1, "5"); SetCell(s, 2, 2, "6");
        });

        var mergeResult = new ThreeWayDiffResult();
        mergeResult.AddCell(new CellMergeResult(1, 0, "1", "10", "1"));  // MineOnly
        mergeResult.AddCell(new CellMergeResult(1, 2, "3", "3", "30")); // TheirsOnly
        mergeResult.AddCell(new CellMergeResult(2, 1, "5", "50", "50")); // BothSame

        MergeWriter.Write(basePath, outputPath, mergeResult, "Sheet1");

        var outWb = ExcelWorkbook.Create(outputPath, new ExcelSheetReadConfig());
        var sheet = outWb.Sheets["Sheet1"];

        Assert.AreEqual("10", sheet.Rows[1].Cells[0].Value, "MineOnly change applied");
        Assert.AreEqual("2", sheet.Rows[1].Cells[1].Value, "Unchanged cell preserved");
        Assert.AreEqual("30", sheet.Rows[1].Cells[2].Value, "TheirsOnly change applied");
        Assert.AreEqual("50", sheet.Rows[2].Cells[1].Value, "BothSame change applied");
    }

    // ═══════════════════════════════════════════════════════════
    // 5. ExcelWorkbook edge cases
    // ═══════════════════════════════════════════════════════════

    [TestMethod]
    public void ExcelWorkbook_MultipleSheets_AllLoaded()
    {
        var path = TestPath("multi_sheet.xlsx");
        CreateMultiSheetExcel(path,
            ("Sales", s =>
            {
                SetCell(s, 0, 0, "Product"); SetCell(s, 0, 1, "Revenue");
                SetCell(s, 1, 0, "Widget");  SetCell(s, 1, 1, "1000");
            }),
            ("Inventory", s =>
            {
                SetCell(s, 0, 0, "Item"); SetCell(s, 0, 1, "Qty");
                SetCell(s, 1, 0, "Widget"); SetCell(s, 1, 1, "50");
                SetCell(s, 2, 0, "Gadget"); SetCell(s, 2, 1, "30");
            }),
            ("Config", s =>
            {
                SetCell(s, 0, 0, "Key"); SetCell(s, 0, 1, "Value");
                SetCell(s, 1, 0, "Version"); SetCell(s, 1, 1, "1.0");
            })
        );

        var wb = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());

        Assert.AreEqual(3, wb.Sheets.Count, "All 3 sheets loaded");
        Assert.IsTrue(wb.Sheets.ContainsKey("Sales"), "Sales sheet present");
        Assert.IsTrue(wb.Sheets.ContainsKey("Inventory"), "Inventory sheet present");
        Assert.IsTrue(wb.Sheets.ContainsKey("Config"), "Config sheet present");

        // Verify content of each sheet
        Assert.AreEqual(2, wb.Sheets["Sales"].Rows.Count, "Sales has 2 rows");
        Assert.AreEqual(3, wb.Sheets["Inventory"].Rows.Count, "Inventory has 3 rows");
        Assert.AreEqual(2, wb.Sheets["Config"].Rows.Count, "Config has 2 rows");

        Assert.AreEqual("Widget", wb.Sheets["Sales"].Rows[1].Cells[0].Value);
        Assert.AreEqual("30", wb.Sheets["Inventory"].Rows[2].Cells[1].Value);
        Assert.AreEqual("1.0", wb.Sheets["Config"].Rows[1].Cells[1].Value);
    }

    [TestMethod]
    public void ExcelWorkbook_GetSheetNames_ReturnsAll()
    {
        var path = TestPath("sheet_names.xlsx");
        CreateMultiSheetExcel(path,
            ("Alpha", s => SetCell(s, 0, 0, "A")),
            ("Beta", s => SetCell(s, 0, 0, "B")),
            ("Gamma", s => SetCell(s, 0, 0, "C"))
        );

        var names = ExcelWorkbook.GetSheetNames(path).ToList();

        Assert.AreEqual(3, names.Count, "3 sheet names");
        CollectionAssert.Contains(names, "Alpha");
        CollectionAssert.Contains(names, "Beta");
        CollectionAssert.Contains(names, "Gamma");
    }

    [TestMethod]
    public void ExcelWorkbook_XlsFormat_ReadSuccessfully()
    {
        var path = TestPath("legacy.xls");
        CreateXlsExcel(path, s =>
        {
            SetCell(s, 0, 0, "ID"); SetCell(s, 0, 1, "Name");
            SetCell(s, 1, 0, "1");  SetCell(s, 1, 1, "Alice");
            SetCell(s, 2, 0, "2");  SetCell(s, 2, 1, "Bob");
        });

        var wb = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());

        Assert.IsNotNull(wb, "XLS workbook loaded");
        Assert.AreEqual(1, wb.Sheets.Count, "1 sheet loaded");
        Assert.IsTrue(wb.Sheets.ContainsKey("Sheet1"), "Sheet1 present");
        Assert.AreEqual(3, wb.Sheets["Sheet1"].Rows.Count, "3 rows loaded");
        Assert.AreEqual("Alice", wb.Sheets["Sheet1"].Rows[1].Cells[1].Value, "Cell value correct");
    }

    [TestMethod]
    public void ExcelWorkbook_XlsVsXlsx_SameContent()
    {
        var xlsPath = TestPath("compare.xls");
        var xlsxPath = TestPath("compare.xlsx");

        Action<ISheet> populate = s =>
        {
            SetCell(s, 0, 0, "Header1"); SetCell(s, 0, 1, "Header2");
            SetCell(s, 1, 0, "Data1");   SetCell(s, 1, 1, "Data2");
        };

        CreateXlsExcel(xlsPath, populate);
        CreateExcel(xlsxPath, populate);

        var xlsWb = ExcelWorkbook.Create(xlsPath, new ExcelSheetReadConfig());
        var xlsxWb = ExcelWorkbook.Create(xlsxPath, new ExcelSheetReadConfig());

        var xlsSheet = xlsWb.Sheets["Sheet1"];
        var xlsxSheet = xlsxWb.Sheets["Sheet1"];

        Assert.AreEqual(xlsSheet.Rows.Count, xlsxSheet.Rows.Count, "Same row count");

        foreach (var kvp in xlsSheet.Rows)
        {
            var xlsRow = kvp.Value;
            var xlsxRow = xlsxSheet.Rows[kvp.Key];
            Assert.AreEqual(xlsRow.Cells.Count, xlsxRow.Cells.Count, $"Row {kvp.Key}: same cell count");
            for (int c = 0; c < xlsRow.Cells.Count; c++)
            {
                Assert.AreEqual(xlsRow.Cells[c].Value, xlsxRow.Cells[c].Value,
                    $"Cell [{kvp.Key},{c}]: same value in xls and xlsx");
            }
        }
    }

    [TestMethod]
    public void ExcelWorkbook_EmptySheet_HandledGracefully()
    {
        var path = TestPath("empty_sheet.xlsx");
        CreateExcel(path, s => { /* no cells */ });

        var wb = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());

        Assert.IsNotNull(wb);
        Assert.AreEqual(1, wb.Sheets.Count, "Sheet exists");
        Assert.AreEqual(0, wb.Sheets["Sheet1"].Rows.Count, "Empty sheet has 0 rows");
    }

    [TestMethod]
    public void ExcelWorkbook_EmptySheetAmongPopulated()
    {
        var path = TestPath("mixed_empty.xlsx");
        CreateMultiSheetExcel(path,
            ("Data", s =>
            {
                SetCell(s, 0, 0, "Hello");
                SetCell(s, 1, 0, "World");
            }),
            ("Empty", s => { /* empty */ }),
            ("MoreData", s =>
            {
                SetCell(s, 0, 0, "Test");
            })
        );

        var wb = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());

        Assert.AreEqual(3, wb.Sheets.Count, "All 3 sheets loaded");
        Assert.AreEqual(2, wb.Sheets["Data"].Rows.Count, "Data sheet has rows");
        Assert.AreEqual(0, wb.Sheets["Empty"].Rows.Count, "Empty sheet has 0 rows");
        Assert.AreEqual(1, wb.Sheets["MoreData"].Rows.Count, "MoreData has rows");
    }

    [TestMethod]
    public void ExcelWorkbook_CsvFile_LoadsAsSheet()
    {
        var path = TestPath("test.csv");
        File.WriteAllText(path, "ID,Name,Score\n1,Alice,90\n2,Bob,85\n");

        var wb = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());

        Assert.AreEqual(1, wb.Sheets.Count, "CSV loads as 1 sheet");
        var sheetName = wb.Sheets.Keys.First();
        Assert.AreEqual("test.csv", sheetName, "Sheet named after file");

        var sheet = wb.Sheets[sheetName];
        Assert.IsTrue(sheet.Rows.Count >= 2, $"At least 2 data rows (got {sheet.Rows.Count})");
    }

    [TestMethod]
    public void ExcelWorkbook_TsvFile_LoadsAsSheet()
    {
        var path = TestPath("test.tsv");
        File.WriteAllText(path, "ID\tName\tScore\n1\tAlice\t90\n2\tBob\t85\n");

        var wb = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());

        Assert.AreEqual(1, wb.Sheets.Count, "TSV loads as 1 sheet");
        var sheetName = wb.Sheets.Keys.First();
        Assert.AreEqual("test.tsv", sheetName, "Sheet named after file");
    }

    // ═══════════════════════════════════════════════════════════
    // 6. CellMergeResult unit tests
    // ═══════════════════════════════════════════════════════════

    [TestMethod]
    public void CellMergeResult_StatusDetermination_AllCases()
    {
        // Unchanged
        var unchanged = new CellMergeResult(0, 0, "val", "val", "val");
        Assert.AreEqual(CellMergeStatus.Unchanged, unchanged.Status);
        Assert.AreEqual("val", unchanged.ResolvedValue);

        // MineOnly
        var mineOnly = new CellMergeResult(0, 0, "base", "mine", "base");
        Assert.AreEqual(CellMergeStatus.MineOnly, mineOnly.Status);
        Assert.AreEqual("mine", mineOnly.ResolvedValue);

        // TheirsOnly
        var theirsOnly = new CellMergeResult(0, 0, "base", "base", "theirs");
        Assert.AreEqual(CellMergeStatus.TheirsOnly, theirsOnly.Status);
        Assert.AreEqual("theirs", theirsOnly.ResolvedValue);

        // BothSame
        var bothSame = new CellMergeResult(0, 0, "base", "same", "same");
        Assert.AreEqual(CellMergeStatus.BothSame, bothSame.Status);
        Assert.AreEqual("same", bothSame.ResolvedValue);

        // Conflict
        var conflict = new CellMergeResult(0, 0, "base", "mine", "theirs");
        Assert.AreEqual(CellMergeStatus.Conflict, conflict.Status);
        Assert.IsNull(conflict.ResolvedValue);
    }

    [TestMethod]
    public void CellMergeResult_NullValues_HandledSafely()
    {
        // null values should be converted to empty strings
        var cell = new CellMergeResult(0, 0, null, null, null);
        Assert.AreEqual(string.Empty, cell.BaseValue);
        Assert.AreEqual(string.Empty, cell.MineValue);
        Assert.AreEqual(string.Empty, cell.TheirsValue);
        Assert.AreEqual(CellMergeStatus.Unchanged, cell.Status);
    }

    [TestMethod]
    public void ThreeWayDiffResult_GetCell_ReturnsCorrectCell()
    {
        var result = new ThreeWayDiffResult();
        result.AddCell(new CellMergeResult(0, 0, "a", "b", "a"));
        result.AddCell(new CellMergeResult(0, 1, "x", "x", "x"));
        result.AddCell(new CellMergeResult(1, 0, "p", "p", "q"));

        var cell00 = result.GetCell(0, 0);
        Assert.IsNotNull(cell00);
        Assert.AreEqual(CellMergeStatus.MineOnly, cell00.Status);

        var cell01 = result.GetCell(0, 1);
        Assert.IsNotNull(cell01);
        Assert.AreEqual(CellMergeStatus.Unchanged, cell01.Status);

        var cell10 = result.GetCell(1, 0);
        Assert.IsNotNull(cell10);
        Assert.AreEqual(CellMergeStatus.TheirsOnly, cell10.Status);

        // Non-existent cell
        var cellNull = result.GetCell(5, 5);
        Assert.IsNull(cellNull, "Non-existent cell returns null");
    }

    [TestMethod]
    public void ThreeWayDiffResult_Counters_Accurate()
    {
        var result = new ThreeWayDiffResult();
        result.AddCell(new CellMergeResult(0, 0, "a", "a", "a")); // Unchanged
        result.AddCell(new CellMergeResult(0, 1, "a", "b", "a")); // MineOnly
        result.AddCell(new CellMergeResult(0, 2, "a", "a", "c")); // TheirsOnly
        result.AddCell(new CellMergeResult(1, 0, "a", "b", "b")); // BothSame
        result.AddCell(new CellMergeResult(1, 1, "a", "b", "c")); // Conflict

        Assert.AreEqual(1, result.ConflictCount, "1 conflict");
        Assert.AreEqual(3, result.AutoMergedCount, "3 auto-merged (MineOnly + TheirsOnly + BothSame)");
        Assert.AreEqual(4, result.TotalChangedCount, "4 changed (all except Unchanged)");
        Assert.IsTrue(result.HasConflicts);
        Assert.AreEqual(1, result.UnresolvedConflictCount, "1 unresolved");
    }

    // ═══════════════════════════════════════════════════════════
    // 7. Integration: 3-way diff + MergeWriter roundtrip
    // ═══════════════════════════════════════════════════════════

    [TestMethod]
    public void Integration_ThreeWayDiffThenMergeWrite_Roundtrip()
    {
        var basePath = TestPath("rt_base.xlsx");
        var minePath = TestPath("rt_mine.xlsx");
        var theirsPath = TestPath("rt_theirs.xlsx");
        var outputPath = TestPath("rt_output.xlsx");

        CreateExcel(basePath, s =>
        {
            SetCell(s, 0, 0, "ID"); SetCell(s, 0, 1, "Name"); SetCell(s, 0, 2, "Score");
            SetCell(s, 1, 0, "1");  SetCell(s, 1, 1, "Alice"); SetCell(s, 1, 2, "90");
            SetCell(s, 2, 0, "2");  SetCell(s, 2, 1, "Bob");   SetCell(s, 2, 2, "85");
            SetCell(s, 3, 0, "3");  SetCell(s, 3, 1, "Carol"); SetCell(s, 3, 2, "88");
        });
        // MINE changes Alice's score
        CreateExcel(minePath, s =>
        {
            SetCell(s, 0, 0, "ID"); SetCell(s, 0, 1, "Name"); SetCell(s, 0, 2, "Score");
            SetCell(s, 1, 0, "1");  SetCell(s, 1, 1, "Alice"); SetCell(s, 1, 2, "95");
            SetCell(s, 2, 0, "2");  SetCell(s, 2, 1, "Bob");   SetCell(s, 2, 2, "85");
            SetCell(s, 3, 0, "3");  SetCell(s, 3, 1, "Carol"); SetCell(s, 3, 2, "88");
        });
        // THEIRS changes Bob's name
        CreateExcel(theirsPath, s =>
        {
            SetCell(s, 0, 0, "ID"); SetCell(s, 0, 1, "Name"); SetCell(s, 0, 2, "Score");
            SetCell(s, 1, 0, "1");  SetCell(s, 1, 1, "Alice"); SetCell(s, 1, 2, "90");
            SetCell(s, 2, 0, "2");  SetCell(s, 2, 1, "Bobby"); SetCell(s, 2, 2, "85");
            SetCell(s, 3, 0, "3");  SetCell(s, 3, 1, "Carol"); SetCell(s, 3, 2, "88");
        });

        var result = ComputeThreeWay(basePath, minePath, theirsPath);

        Assert.AreEqual(0, result.ConflictCount, "No conflicts for non-overlapping changes");

        // Write merged result
        MergeWriter.Write(basePath, outputPath, result, "Sheet1");

        // Read and verify
        var outWb = ExcelWorkbook.Create(outputPath, new ExcelSheetReadConfig());
        var outSheet = outWb.Sheets["Sheet1"];

        // Alice's score should be 95 (from MINE)
        Assert.AreEqual("95", outSheet.Rows[1].Cells[2].Value, "MINE change (Alice score 95) applied");

        // Bob's name should be Bobby (from THEIRS)
        Assert.AreEqual("Bobby", outSheet.Rows[2].Cells[1].Value, "THEIRS change (Bobby) applied");

        // Carol unchanged
        Assert.AreEqual("Carol", outSheet.Rows[3].Cells[1].Value, "Unchanged row preserved");
        Assert.AreEqual("88", outSheet.Rows[3].Cells[2].Value, "Unchanged cell preserved");
    }

    // ═══════════════════════════════════════════════════════════
    // 8. Diff with various configurations
    // ═══════════════════════════════════════════════════════════

    [TestMethod]
    public void ExcelSheet_Diff_FormulaAndWhitespace_Combined()
    {
        var src = TestPath("combo_src.xlsx");
        var dst = TestPath("combo_dst.xlsx");

        CreateExcel(src, s =>
        {
            SetCell(s, 0, 0, "  Hello  ");
            SetFormula(s, 0, 1, "A1");
        });
        CreateExcel(dst, s =>
        {
            SetCell(s, 0, 0, "Hello");
            SetFormula(s, 0, 1, "B1"); // different formula
        });

        var config = new ExcelSheetReadConfig();
        var srcWb = ExcelWorkbook.Create(src, config);
        var dstWb = ExcelWorkbook.Create(dst, config);

        // With ignore whitespace + compare formula
        var diff = ExcelSheet.Diff(srcWb.Sheets["Sheet1"], dstWb.Sheets["Sheet1"],
            new ExcelSheetDiffConfig { IgnoreWhitespace = true, CompareFormula = true });
        var summary = diff.CreateSummary();

        // Whitespace diff ignored but formula diff detected
        Assert.IsTrue(summary.HasDiff, "Formula diff detected even with whitespace ignored");
        Assert.IsTrue(summary.ModifiedCellCount >= 1, "At least 1 modified cell for formula");
    }

    [TestMethod]
    public void ExcelSheet_Diff_TwoEmptySheets_NoDiff()
    {
        var src = TestPath("empty1.xlsx");
        var dst = TestPath("empty2.xlsx");

        CreateExcel(src, s => { });
        CreateExcel(dst, s => { });

        var config = new ExcelSheetReadConfig();
        var srcWb = ExcelWorkbook.Create(src, config);
        var dstWb = ExcelWorkbook.Create(dst, config);

        var diff = ExcelSheet.Diff(srcWb.Sheets["Sheet1"], dstWb.Sheets["Sheet1"],
            new ExcelSheetDiffConfig());
        var summary = diff.CreateSummary();

        Assert.IsFalse(summary.HasDiff, "Two empty sheets have no diff");
    }
}
