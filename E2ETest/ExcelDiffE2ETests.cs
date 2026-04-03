using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using ExcelMerge;

namespace E2ETest;

[TestClass]
public class ExcelDiffE2ETests
{
    private string _testDir;

    [TestInitialize]
    public void Setup()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"ExcelMerge_E2E_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);
    }

    [TestCleanup]
    public void Cleanup()
    {
        try { Directory.Delete(_testDir, true); } catch { }
    }

    #region Helpers

    private string TestPath(string name) => Path.Combine(_testDir, name);

    private void CreateExcel(string path, Action<ISheet> populate)
    {
        var wb = new XSSFWorkbook();
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

    private ExcelSheetDiffSummary DiffFiles(string src, string dst, bool compareFormula = false)
    {
        var config = new ExcelSheetReadConfig();
        var srcWb = ExcelWorkbook.Create(src, config);
        var dstWb = ExcelWorkbook.Create(dst, config);
        var diff = ExcelSheet.Diff(
            srcWb.Sheets["Sheet1"], dstWb.Sheets["Sheet1"],
            new ExcelSheetDiffConfig { CompareFormula = compareFormula });
        return diff.CreateSummary();
    }

    #endregion

    // ── Test 1: Identical files → no diff ──
    [TestMethod]
    public void IdenticalFiles_NoDiff()
    {
        var src = TestPath("identical_src.xlsx");
        CreateExcel(src, s =>
        {
            SetCell(s, 0, 0, "ID"); SetCell(s, 0, 1, "Name");
            SetCell(s, 1, 0, "1"); SetCell(s, 1, 1, "Alice");
            SetCell(s, 2, 0, "2"); SetCell(s, 2, 1, "Bob");
        });
        File.Copy(src, TestPath("identical_dst.xlsx"));

        var summary = DiffFiles(src, TestPath("identical_dst.xlsx"));

        Assert.AreEqual(0, summary.ModifiedCellCount, "Modified cells");
        Assert.AreEqual(0, summary.AddedRowCount, "Added rows");
        Assert.AreEqual(0, summary.RemovedRowCount, "Removed rows");
        Assert.IsFalse(summary.HasDiff);
    }

    // ── Test 2: Cell modifications detected ──
    [TestMethod]
    public void CellModification_Detected()
    {
        var src = TestPath("mod_src.xlsx");
        var dst = TestPath("mod_dst.xlsx");
        CreateExcel(src, s =>
        {
            SetCell(s, 0, 0, "1"); SetCell(s, 0, 1, "Alice"); SetCell(s, 0, 2, "90");
            SetCell(s, 1, 0, "2"); SetCell(s, 1, 1, "Bob");   SetCell(s, 1, 2, "85");
        });
        CreateExcel(dst, s =>
        {
            SetCell(s, 0, 0, "1"); SetCell(s, 0, 1, "Alice"); SetCell(s, 0, 2, "95");  // changed
            SetCell(s, 1, 0, "2"); SetCell(s, 1, 1, "Bobby"); SetCell(s, 1, 2, "85");  // changed
        });

        var summary = DiffFiles(src, dst);

        Assert.AreEqual(2, summary.ModifiedCellCount);
        Assert.AreEqual(2, summary.ModifiedRowCount);
        Assert.IsTrue(summary.HasDiff);
    }

    // ── Test 3: Row insertion — detects added row ──
    [TestMethod]
    public void RowInsertion_DetectsAddedRow()
    {
        var src = TestPath("insert_src.xlsx");
        var dst = TestPath("insert_dst.xlsx");
        CreateExcel(src, s =>
        {
            SetCell(s, 0, 0, "A"); SetCell(s, 0, 1, "1");
            SetCell(s, 1, 0, "B"); SetCell(s, 1, 1, "2");
            SetCell(s, 2, 0, "C"); SetCell(s, 2, 1, "3");
        });
        CreateExcel(dst, s =>
        {
            SetCell(s, 0, 0, "A"); SetCell(s, 0, 1, "1");
            SetCell(s, 1, 0, "X"); SetCell(s, 1, 1, "9");  // inserted
            SetCell(s, 2, 0, "B"); SetCell(s, 2, 1, "2");
            SetCell(s, 3, 0, "C"); SetCell(s, 3, 1, "3");
        });

        var summary = DiffFiles(src, dst);

        // LCS OptimizeCaseDeletedFirst may fold adjacent Delete+Insert into Modified
        Assert.IsTrue(summary.HasDiff, "Diff detected");
        var totalChangedRows = summary.AddedRowCount + summary.ModifiedRowCount;
        Assert.IsTrue(totalChangedRows >= 1, $"At least 1 changed row (got added={summary.AddedRowCount}, modified={summary.ModifiedRowCount})");
        // Key check: unchanged rows (A, B, C) should not ALL be marked as modified
        Assert.IsTrue(summary.ModifiedRowCount <= 1, $"At most 1 false modification from LCS folding (got {summary.ModifiedRowCount})");
    }

    // ── Test 4: Row deletion — detects removed row ──
    [TestMethod]
    public void RowDeletion_DetectsRemovedRow()
    {
        var src = TestPath("del_src.xlsx");
        var dst = TestPath("del_dst.xlsx");
        CreateExcel(src, s =>
        {
            SetCell(s, 0, 0, "A"); SetCell(s, 0, 1, "1");
            SetCell(s, 1, 0, "B"); SetCell(s, 1, 1, "2");
            SetCell(s, 2, 0, "C"); SetCell(s, 2, 1, "3");
            SetCell(s, 3, 0, "D"); SetCell(s, 3, 1, "4");
        });
        CreateExcel(dst, s =>
        {
            SetCell(s, 0, 0, "A"); SetCell(s, 0, 1, "1");
            SetCell(s, 1, 0, "C"); SetCell(s, 1, 1, "3");  // B deleted
            SetCell(s, 2, 0, "D"); SetCell(s, 2, 1, "4");
        });

        var summary = DiffFiles(src, dst);

        Assert.IsTrue(summary.HasDiff, "Diff detected");
        var totalChangedRows = summary.RemovedRowCount + summary.ModifiedRowCount;
        Assert.IsTrue(totalChangedRows >= 1, $"At least 1 changed row (got removed={summary.RemovedRowCount}, modified={summary.ModifiedRowCount})");
        Assert.IsTrue(summary.ModifiedRowCount <= 1, $"At most 1 false modification from LCS folding (got {summary.ModifiedRowCount})");
    }

    // ── Test 5: Formula diff — detects formula changes ──
    [TestMethod]
    public void FormulaDiff_DetectsFormulaChange()
    {
        var src = TestPath("formula_src.xlsx");
        var dst = TestPath("formula_dst.xlsx");
        CreateExcel(src, s =>
        {
            SetCell(s, 0, 0, "Value"); SetCell(s, 0, 1, "Total");
            SetNumeric(s, 1, 0, 10);
            SetFormula(s, 1, 1, "A2*2");
        });
        CreateExcel(dst, s =>
        {
            SetCell(s, 0, 0, "Value"); SetCell(s, 0, 1, "Total");
            SetNumeric(s, 1, 0, 10);
            SetFormula(s, 1, 1, "A2*3");  // formula changed
        });

        var summary = DiffFiles(src, dst, compareFormula: true);

        Assert.IsTrue(summary.ModifiedCellCount >= 1, $"Formula mode should detect change, got {summary.ModifiedCellCount}");
    }

    // ── Test 6: Formula property is populated ──
    [TestMethod]
    public void FormulaProperty_IsPopulated()
    {
        var path = TestPath("formula_read.xlsx");
        CreateExcel(path, s =>
        {
            SetNumeric(s, 0, 0, 10);
            SetFormula(s, 0, 1, "A1*2");
            SetCell(s, 0, 2, "plain text");
        });

        var wb = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());
        var row = wb.Sheets["Sheet1"].Rows.Values.First();

        Assert.AreEqual(string.Empty, row.Cells[0].Formula, "Numeric cell has no formula");
        Assert.AreEqual("=A1*2", row.Cells[1].Formula, "Formula cell has formula");
        Assert.AreEqual(string.Empty, row.Cells[2].Formula, "Text cell has no formula");
    }

    // ── Test 7: File locking — can read file locked by another process ──
    [TestMethod]
    public void FileLocking_CanReadLockedFile()
    {
        var path = TestPath("locked.xlsx");
        CreateExcel(path, s =>
        {
            SetCell(s, 0, 0, "Test"); SetCell(s, 0, 1, "Data");
        });

        // Simulate lock by another process
        using var lockStream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

        var wb = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());

        Assert.IsNotNull(wb);
        Assert.AreEqual(1, wb.Sheets.Count);
    }

    // ── Test 8: Large file — 200 rows with 3 inserted in middle ──
    [TestMethod]
    public void LargeFile_RowInsertionAligned()
    {
        var src = TestPath("large_src.xlsx");
        var dst = TestPath("large_dst.xlsx");
        CreateExcel(src, s =>
        {
            for (int i = 0; i < 200; i++)
            {
                SetCell(s, i, 0, $"Row{i}");
                SetCell(s, i, 1, $"Data{i}");
            }
        });
        CreateExcel(dst, s =>
        {
            for (int i = 0; i < 100; i++)
            {
                SetCell(s, i, 0, $"Row{i}");
                SetCell(s, i, 1, $"Data{i}");
            }
            SetCell(s, 100, 0, "NEW1"); SetCell(s, 100, 1, "Inserted");
            SetCell(s, 101, 0, "NEW2"); SetCell(s, 101, 1, "Inserted");
            SetCell(s, 102, 0, "NEW3"); SetCell(s, 102, 1, "Inserted");
            for (int i = 100; i < 200; i++)
            {
                SetCell(s, i + 3, 0, $"Row{i}");
                SetCell(s, i + 3, 1, $"Data{i}");
            }
        });

        var summary = DiffFiles(src, dst);

        var totalChanged = summary.AddedRowCount + summary.ModifiedRowCount;
        Assert.IsTrue(totalChanged >= 3, $"At least 3 changed rows for 3 insertions (got added={summary.AddedRowCount}, modified={summary.ModifiedRowCount})");
        // Key: the 200 unchanged rows should not all be marked as modified
        Assert.IsTrue(summary.ModifiedRowCount <= 3, $"At most 3 modified from LCS folding (got {summary.ModifiedRowCount})");
    }

    // ── Test 9: Empty file vs populated file ──
    [TestMethod]
    public void EmptyFile_DiffWithPopulated()
    {
        var src = TestPath("empty.xlsx");
        var dst = TestPath("populated.xlsx");
        CreateExcel(src, s => { });
        CreateExcel(dst, s =>
        {
            SetCell(s, 0, 0, "New"); SetCell(s, 0, 1, "Data");
        });

        var summary = DiffFiles(src, dst);

        Assert.IsTrue(summary.AddedRowCount >= 1, "Rows added when src is empty");
    }

    // ── Test 10: CSV file locking ──
    [TestMethod]
    public void CsvFile_CanReadLocked()
    {
        var path = TestPath("test.csv");
        File.WriteAllText(path, "ID,Name\n1,Alice\n2,Bob\n");

        using var lockStream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

        var wb = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());
        Assert.IsNotNull(wb);
        Assert.AreEqual(1, wb.Sheets.Count);
    }
}
