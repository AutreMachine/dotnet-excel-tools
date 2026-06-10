using ExcelTool.Xlsx;
using Xunit;

namespace ExcelTool.Tests;

public class WorkbookTests
{
    [Fact]
    public void NewWorkbook_HasZeroSheets()
    {
        var wb = new XSSFWorkbook();
        Assert.Equal(0, wb.NumberOfSheets);
    }

    [Fact]
    public void NewWorkbook_CanWriteToStream()
    {
        var wb = new XSSFWorkbook();
        wb.CreateSheet("Data");
        using var ms = new MemoryStream();
        wb.Write(ms);
        Assert.True(ms.Length > 0);
    }

    [Fact]
    public void WriteAndReopen_PreservesSheetCount()
    {
        using var ms = new MemoryStream();
        using (var wb = new XSSFWorkbook())
        {
            wb.CreateSheet("Alpha");
            wb.CreateSheet("Beta");
            wb.Write(ms);
        }
        ms.Position = 0;
        var loaded = new XSSFWorkbook(ms);
        Assert.Equal(2, loaded.NumberOfSheets);
    }

    [Fact]
    public void WriteAndReopen_PreservesSheetNames()
    {
        using var ms = new MemoryStream();
        using (var wb = new XSSFWorkbook())
        {
            wb.CreateSheet("Alpha");
            wb.CreateSheet("Beta");
            wb.Write(ms);
        }
        ms.Position = 0;
        var loaded = new XSSFWorkbook(ms);
        Assert.Equal("Alpha", loaded.GetSheetAt(0).SheetName);
        Assert.Equal("Beta", loaded.GetSheetAt(1).SheetName);
    }

    [Fact]
    public void WriteAndReopen_PreservesStringCellValue()
    {
        using var ms = new MemoryStream();
        using (var wb = new XSSFWorkbook())
        {
            var sheet = wb.CreateSheet("Sheet1");
            var row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("Hello World");
            wb.Write(ms);
        }
        ms.Position = 0;
        var loaded = new XSSFWorkbook(ms);
        Assert.Equal("Hello World", loaded.GetSheetAt(0).GetRow(0)!.GetCell(0)!.StringCellValue);
    }

    [Fact]
    public void WriteAndReopen_PreservesNumericCellValue()
    {
        using var ms = new MemoryStream();
        using (var wb = new XSSFWorkbook())
        {
            var sheet = wb.CreateSheet("Sheet1");
            var row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue(42.5);
            wb.Write(ms);
        }
        ms.Position = 0;
        var loaded = new XSSFWorkbook(ms);
        Assert.Equal(42.5, loaded.GetSheetAt(0).GetRow(0)!.GetCell(0)!.NumericCellValue);
    }

    [Fact]
    public void WriteAndReopen_PreservesBooleanCellValue()
    {
        using var ms = new MemoryStream();
        using (var wb = new XSSFWorkbook())
        {
            var sheet = wb.CreateSheet("Sheet1");
            var row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue(true);
            wb.Write(ms);
        }
        ms.Position = 0;
        var loaded = new XSSFWorkbook(ms);
        Assert.True(loaded.GetSheetAt(0).GetRow(0)!.GetCell(0)!.BooleanCellValue);
    }

    [Fact]
    public void OpenStaticFactory_LoadsSameAsConstructor()
    {
        using var ms = new MemoryStream();
        using (var wb = new XSSFWorkbook())
        {
            wb.CreateSheet("Test");
            wb.Write(ms);
        }
        ms.Position = 0;
        var loaded = XSSFWorkbook.Open(ms);
        Assert.Equal(1, loaded.NumberOfSheets);
    }

    [Fact]
    public void Dispose_DoesNotThrow()
    {
        var wb = new XSSFWorkbook();
        wb.Dispose();
        wb.Dispose(); // double-dispose should also be safe
    }
}
