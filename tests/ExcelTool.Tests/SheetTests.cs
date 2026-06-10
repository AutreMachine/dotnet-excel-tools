using ExcelTool.Xlsx;
using Xunit;

namespace ExcelTool.Tests;

public class SheetTests
{
    private static XSSFWorkbook CreateWorkbook() => new();

    [Fact]
    public void CreateSheet_AutoName_IncreasesSheetCount()
    {
        var wb = CreateWorkbook();
        wb.CreateSheet();
        Assert.Equal(1, wb.NumberOfSheets);
    }

    [Fact]
    public void CreateSheet_AutoName_UsesSheet1Sheet2Convention()
    {
        var wb = CreateWorkbook();
        var s1 = wb.CreateSheet();
        var s2 = wb.CreateSheet();
        Assert.Equal("Sheet1", s1.SheetName);
        Assert.Equal("Sheet2", s2.SheetName);
    }

    [Fact]
    public void CreateSheet_WithName_HasCorrectName()
    {
        var wb = CreateWorkbook();
        var sheet = wb.CreateSheet("MySheet");
        Assert.Equal("MySheet", sheet.SheetName);
    }

    [Fact]
    public void CreateMultipleSheets_AllAccessibleByIndex()
    {
        var wb = CreateWorkbook();
        wb.CreateSheet("A");
        wb.CreateSheet("B");
        wb.CreateSheet("C");
        Assert.Equal("A", wb.GetSheetAt(0).SheetName);
        Assert.Equal("B", wb.GetSheetAt(1).SheetName);
        Assert.Equal("C", wb.GetSheetAt(2).SheetName);
    }

    [Fact]
    public void CreateMultipleSheets_CountIsCorrect()
    {
        var wb = CreateWorkbook();
        wb.CreateSheet("A");
        wb.CreateSheet("B");
        wb.CreateSheet("C");
        Assert.Equal(3, wb.NumberOfSheets);
    }

    [Fact]
    public void GetSheet_ByName_ReturnsCorrectSheet()
    {
        var wb = CreateWorkbook();
        wb.CreateSheet("Alpha");
        wb.CreateSheet("Beta");
        Assert.Equal("Beta", wb.GetSheet("Beta")!.SheetName);
    }

    [Fact]
    public void GetSheet_ByName_IsCaseInsensitive()
    {
        var wb = CreateWorkbook();
        wb.CreateSheet("Alpha");
        Assert.NotNull(wb.GetSheet("alpha"));
        Assert.NotNull(wb.GetSheet("ALPHA"));
    }

    [Fact]
    public void GetSheet_UnknownName_ReturnsNull()
    {
        var wb = CreateWorkbook();
        Assert.Null(wb.GetSheet("DoesNotExist"));
    }

    [Fact]
    public void GetSheetIndex_ReturnsCorrectIndex()
    {
        var wb = CreateWorkbook();
        wb.CreateSheet("First");
        wb.CreateSheet("Second");
        wb.CreateSheet("Third");
        Assert.Equal(0, wb.GetSheetIndex("First"));
        Assert.Equal(1, wb.GetSheetIndex("Second"));
        Assert.Equal(2, wb.GetSheetIndex("Third"));
    }

    [Fact]
    public void GetSheetIndex_UnknownName_ReturnsMinusOne()
    {
        var wb = CreateWorkbook();
        Assert.Equal(-1, wb.GetSheetIndex("NoSuch"));
    }

    [Fact]
    public void CreateSheet_DuplicateName_Throws()
    {
        var wb = CreateWorkbook();
        wb.CreateSheet("Dup");
        Assert.Throws<ArgumentException>(() => wb.CreateSheet("Dup"));
    }

    [Fact]
    public void RemoveSheetAt_DecreasesCount()
    {
        var wb = CreateWorkbook();
        wb.CreateSheet("A");
        wb.CreateSheet("B");
        wb.RemoveSheetAt(0);
        Assert.Equal(1, wb.NumberOfSheets);
    }

    [Fact]
    public void RemoveSheetAt_ShiftsSurvivingSheets()
    {
        var wb = CreateWorkbook();
        wb.CreateSheet("A");
        wb.CreateSheet("B");
        wb.CreateSheet("C");
        wb.RemoveSheetAt(1); // remove "B"
        Assert.Equal("A", wb.GetSheetAt(0).SheetName);
        Assert.Equal("C", wb.GetSheetAt(1).SheetName);
    }

    [Fact]
    public void RemoveSheetAt_RemovedSheetNoLongerReturnedByName()
    {
        var wb = CreateWorkbook();
        wb.CreateSheet("Removed");
        wb.CreateSheet("Remaining");
        wb.RemoveSheetAt(0);
        Assert.Null(wb.GetSheet("Removed"));
    }

    [Fact]
    public void RemoveAllSheets_WorkbookIsEmpty()
    {
        var wb = CreateWorkbook();
        wb.CreateSheet("A");
        wb.CreateSheet("B");
        wb.RemoveSheetAt(0);
        wb.RemoveSheetAt(0);
        Assert.Equal(0, wb.NumberOfSheets);
    }

    [Fact]
    public void GetSheetAt_OutOfRange_Throws()
    {
        var wb = CreateWorkbook();
        Assert.Throws<ArgumentOutOfRangeException>(() => wb.GetSheetAt(0));
    }

    [Fact]
    public void RemoveSheetAt_OutOfRange_Throws()
    {
        var wb = CreateWorkbook();
        Assert.Throws<ArgumentOutOfRangeException>(() => wb.RemoveSheetAt(0));
    }

    [Fact]
    public void WriteAndReopen_MultipleSheets_Preserved()
    {
        using var ms = new MemoryStream();
        using (var wb = new XSSFWorkbook())
        {
            wb.CreateSheet("First");
            wb.CreateSheet("Second");
            wb.CreateSheet("Third");
            wb.Write(ms);
        }
        ms.Position = 0;
        var loaded = new XSSFWorkbook(ms);
        Assert.Equal(3, loaded.NumberOfSheets);
        Assert.Equal("First", loaded.GetSheetAt(0).SheetName);
        Assert.Equal("Second", loaded.GetSheetAt(1).SheetName);
        Assert.Equal("Third", loaded.GetSheetAt(2).SheetName);
    }
}
