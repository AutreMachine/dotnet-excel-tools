using ExcelTool.Interfaces;
using ExcelTool.Xlsx;
using Xunit;

namespace ExcelTool.Tests;

public class RowTests
{
    private static ISheet CreateSheet()
    {
        var wb = new XSSFWorkbook();
        return wb.CreateSheet("Test");
    }

    [Fact]
    public void CreateRow_IsRetrievableByIndex()
    {
        var sheet = CreateSheet();
        sheet.CreateRow(0);
        Assert.NotNull(sheet.GetRow(0));
    }

    [Fact]
    public void CreateRow_RowNum_MatchesPassedIndex()
    {
        var sheet = CreateSheet();
        var row = sheet.CreateRow(5);
        Assert.Equal(5, row.RowNum);
    }

    [Fact]
    public void CreateMultipleRows_AllRetrievable()
    {
        var sheet = CreateSheet();
        sheet.CreateRow(0);
        sheet.CreateRow(1);
        sheet.CreateRow(2);
        Assert.NotNull(sheet.GetRow(0));
        Assert.NotNull(sheet.GetRow(1));
        Assert.NotNull(sheet.GetRow(2));
    }

    [Fact]
    public void GetRow_NonExistentIndex_ReturnsNull()
    {
        var sheet = CreateSheet();
        Assert.Null(sheet.GetRow(99));
    }

    [Fact]
    public void CreateRow_AtNonContiguousIndices_AllRetrievable()
    {
        var sheet = CreateSheet();
        sheet.CreateRow(0);
        sheet.CreateRow(5);
        sheet.CreateRow(100);
        Assert.NotNull(sheet.GetRow(0));
        Assert.NotNull(sheet.GetRow(5));
        Assert.NotNull(sheet.GetRow(100));
        Assert.Null(sheet.GetRow(1));
    }

    [Fact]
    public void CreateRow_UpdatesFirstRowNum()
    {
        var sheet = CreateSheet();
        sheet.CreateRow(3);
        sheet.CreateRow(7);
        Assert.Equal(3, sheet.FirstRowNum);
    }

    [Fact]
    public void CreateRow_UpdatesLastRowNum()
    {
        var sheet = CreateSheet();
        sheet.CreateRow(3);
        sheet.CreateRow(7);
        Assert.Equal(7, sheet.LastRowNum);
    }

    [Fact]
    public void RemoveRow_RowNoLongerRetrievable()
    {
        var sheet = CreateSheet();
        var row = sheet.CreateRow(0);
        sheet.RemoveRow(row);
        Assert.Null(sheet.GetRow(0));
    }

    [Fact]
    public void RemoveRow_DoesNotAffectOtherRows()
    {
        var sheet = CreateSheet();
        var row0 = sheet.CreateRow(0);
        sheet.CreateRow(1);
        sheet.RemoveRow(row0);
        Assert.NotNull(sheet.GetRow(1));
    }

    [Fact]
    public void RemoveRow_MiddleRow_SurroundingRowsUnaffected()
    {
        var sheet = CreateSheet();
        sheet.CreateRow(0);
        var row1 = sheet.CreateRow(1);
        sheet.CreateRow(2);
        sheet.RemoveRow(row1);
        Assert.NotNull(sheet.GetRow(0));
        Assert.Null(sheet.GetRow(1));
        Assert.NotNull(sheet.GetRow(2));
    }

    [Fact]
    public void Row_InitialFirstCellNum_IsMinusOne()
    {
        var sheet = CreateSheet();
        var row = sheet.CreateRow(0);
        Assert.Equal(-1, row.FirstCellNum);
    }

    [Fact]
    public void Row_InitialLastCellNum_IsMinusOne()
    {
        var sheet = CreateSheet();
        var row = sheet.CreateRow(0);
        Assert.Equal(-1, row.LastCellNum);
    }

    [Fact]
    public void Row_Height_DefaultIsZero()
    {
        var sheet = CreateSheet();
        var row = sheet.CreateRow(0);
        Assert.Equal(0, row.Height);
    }

    [Fact]
    public void Row_Height_CanBeSet()
    {
        var sheet = CreateSheet();
        var row = sheet.CreateRow(0);
        row.Height = 30.0;
        Assert.Equal(30.0, row.Height);
    }

    [Fact]
    public void CreateRow_ReplacesExistingRowAtSameIndex()
    {
        var sheet = CreateSheet();
        var row1 = sheet.CreateRow(0);
        row1.CreateCell(0).SetCellValue("original");
        var row2 = sheet.CreateRow(0); // replace
        Assert.Null(row2.GetCell(0)); // new row has no cells
    }

    [Fact]
    public void WriteAndReopen_RowsPreserved()
    {
        using var ms = new MemoryStream();
        using (var wb = new XSSFWorkbook())
        {
            var sheet = wb.CreateSheet("Sheet1");
            var row0 = sheet.CreateRow(0);
            row0.CreateCell(0).SetCellValue("row0");
            var row2 = sheet.CreateRow(2);
            row2.CreateCell(0).SetCellValue("row2");
            wb.Write(ms);
        }
        ms.Position = 0;
        var loaded = new XSSFWorkbook(ms);
        var s = loaded.GetSheetAt(0);
        Assert.Equal("row0", s.GetRow(0)!.GetCell(0)!.StringCellValue);
        Assert.Null(s.GetRow(1));
        Assert.Equal("row2", s.GetRow(2)!.GetCell(0)!.StringCellValue);
    }
}
