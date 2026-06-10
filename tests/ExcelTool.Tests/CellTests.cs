using ExcelTool.Interfaces;
using ExcelTool.Xlsx;
using Xunit;

namespace ExcelTool.Tests;

public class CellTests
{
    private static IRow CreateRow(int rowIndex = 0)
    {
        var wb = new XSSFWorkbook();
        var sheet = wb.CreateSheet("Test");
        return sheet.CreateRow(rowIndex);
    }

    [Fact]
    public void CreateCell_DefaultType_IsBlank()
    {
        var cell = CreateRow().CreateCell(0);
        Assert.Equal(CellType.Blank, cell.CellType);
    }

    [Fact]
    public void CreateCell_IsRetrievableByColumnIndex()
    {
        var row = CreateRow();
        row.CreateCell(3);
        Assert.NotNull(row.GetCell(3));
    }

    [Fact]
    public void CreateCell_ColumnIndex_MatchesPassedValue()
    {
        var cell = CreateRow().CreateCell(5);
        Assert.Equal(5, cell.ColumnIndex);
    }

    [Fact]
    public void CreateCell_RowIndex_MatchesParentRow()
    {
        var wb = new XSSFWorkbook();
        var sheet = wb.CreateSheet("Test");
        var row = sheet.CreateRow(4);
        var cell = row.CreateCell(0);
        Assert.Equal(4, cell.RowIndex);
    }

    [Fact]
    public void CreateCell_WithExplicitBlankType_IsBlank()
    {
        var cell = CreateRow().CreateCell(0, CellType.Blank);
        Assert.Equal(CellType.Blank, cell.CellType);
    }

    [Fact]
    public void CreateCell_WithStringType_HasStringType()
    {
        var cell = CreateRow().CreateCell(0, CellType.String);
        Assert.Equal(CellType.String, cell.CellType);
    }

    [Fact]
    public void CreateCell_WithNumericType_HasNumericType()
    {
        var cell = CreateRow().CreateCell(0, CellType.Numeric);
        Assert.Equal(CellType.Numeric, cell.CellType);
    }

    // --- String values ---

    [Fact]
    public void SetCellValue_String_SetsTypeAndValue()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue("Hello");
        Assert.Equal(CellType.String, cell.CellType);
        Assert.Equal("Hello", cell.StringCellValue);
    }

    [Fact]
    public void SetCellValue_EmptyString_IsString()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue(string.Empty);
        Assert.Equal(CellType.String, cell.CellType);
        Assert.Equal(string.Empty, cell.StringCellValue);
    }

    [Fact]
    public void SetCellValue_StringWithSpecialChars_PreservesValue()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue("Line1\nLine2\tTab");
        Assert.Equal("Line1\nLine2\tTab", cell.StringCellValue);
    }

    // --- Numeric values ---

    [Fact]
    public void SetCellValue_Double_SetsTypeAndValue()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue(3.14);
        Assert.Equal(CellType.Numeric, cell.CellType);
        Assert.Equal(3.14, cell.NumericCellValue);
    }

    [Fact]
    public void SetCellValue_NegativeDouble_Stored()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue(-99.99);
        Assert.Equal(-99.99, cell.NumericCellValue);
    }

    [Fact]
    public void SetCellValue_Zero_Stored()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue(0.0);
        Assert.Equal(0.0, cell.NumericCellValue);
    }

    [Fact]
    public void SetCellValue_LargeNumber_Stored()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue(1_000_000_000.0);
        Assert.Equal(1_000_000_000.0, cell.NumericCellValue);
    }

    // --- Boolean values ---

    [Fact]
    public void SetCellValue_BoolTrue_SetsTypeAndValue()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue(true);
        Assert.Equal(CellType.Boolean, cell.CellType);
        Assert.True(cell.BooleanCellValue);
    }

    [Fact]
    public void SetCellValue_BoolFalse_StoredCorrectly()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue(false);
        Assert.Equal(CellType.Boolean, cell.CellType);
        Assert.False(cell.BooleanCellValue);
    }

    // --- DateTime values ---

    [Fact]
    public void SetCellValue_DateTime_SetsNumericTypeAndIsDateTimeFlag()
    {
        var cell = CreateRow().CreateCell(0);
        var dt = new DateTime(2024, 6, 15, 12, 0, 0);
        cell.SetCellValue(dt);
        Assert.Equal(CellType.Numeric, cell.CellType);
        Assert.True(cell.IsDateTimeCell);
    }

    [Fact]
    public void SetCellValue_DateTime_DateCellValueReturnsOriginalValue()
    {
        var cell = CreateRow().CreateCell(0);
        var dt = new DateTime(2024, 6, 15, 12, 30, 45);
        cell.SetCellValue(dt);
        Assert.Equal(dt, cell.DateCellValue);
    }

    [Fact]
    public void SetCellValue_DateTime_NumericCellValue_Throws()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue(new DateTime(2024, 1, 1));
        Assert.Throws<InvalidOperationException>(() => _ = cell.NumericCellValue);
    }

    // --- Blank / SetBlank ---

    [Fact]
    public void SetBlank_ResetsToBlankType()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue("Not blank");
        cell.SetBlank();
        Assert.Equal(CellType.Blank, cell.CellType);
    }

    [Fact]
    public void SetCellType_Blank_ClearsValue()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue(42.0);
        cell.SetCellType(CellType.Blank);
        Assert.Equal(CellType.Blank, cell.CellType);
    }

    // --- Wrong-type accessor exceptions ---

    [Fact]
    public void StringCellValue_OnNumericCell_Throws()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue(42.0);
        Assert.Throws<InvalidOperationException>(() => _ = cell.StringCellValue);
    }

    [Fact]
    public void NumericCellValue_OnStringCell_Throws()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue("text");
        Assert.Throws<InvalidOperationException>(() => _ = cell.NumericCellValue);
    }

    [Fact]
    public void BooleanCellValue_OnStringCell_Throws()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue("Not a bool");
        Assert.Throws<InvalidOperationException>(() => _ = cell.BooleanCellValue);
    }

    [Fact]
    public void DateCellValue_OnStringCell_Throws()
    {
        var cell = CreateRow().CreateCell(0);
        cell.SetCellValue("Not a date");
        Assert.Throws<InvalidOperationException>(() => _ = cell.DateCellValue);
    }

    // --- Remove cells ---

    [Fact]
    public void RemoveCell_CellNoLongerRetrievable()
    {
        var row = CreateRow();
        var cell = row.CreateCell(2);
        row.RemoveCell(cell);
        Assert.Null(row.GetCell(2));
    }

    [Fact]
    public void RemoveCell_DoesNotAffectOtherCells()
    {
        var row = CreateRow();
        var cell0 = row.CreateCell(0);
        row.CreateCell(1);
        row.RemoveCell(cell0);
        Assert.Null(row.GetCell(0));
        Assert.NotNull(row.GetCell(1));
    }

    [Fact]
    public void RemoveAllCells_RowHasNoRemainingCells()
    {
        var row = CreateRow();
        var c0 = row.CreateCell(0);
        var c1 = row.CreateCell(1);
        row.RemoveCell(c0);
        row.RemoveCell(c1);
        Assert.Equal(-1, row.FirstCellNum);
        Assert.Equal(-1, row.LastCellNum);
    }

    // --- FirstCellNum / LastCellNum after adding cells ---

    [Fact]
    public void Row_FirstCellNum_AfterAddingCells()
    {
        var row = CreateRow();
        row.CreateCell(3);
        row.CreateCell(7);
        Assert.Equal(3, row.FirstCellNum);
    }

    [Fact]
    public void Row_LastCellNum_IsLastColumnPlusOne()
    {
        var row = CreateRow();
        row.CreateCell(0);
        row.CreateCell(4);
        Assert.Equal(5, row.LastCellNum); // last + 1
    }

    // --- Multiple cells in same row ---

    [Fact]
    public void MultipleCellsInRow_DifferentTypes_AllRetrievable()
    {
        var row = CreateRow();
        row.CreateCell(0).SetCellValue("text");
        row.CreateCell(1).SetCellValue(123.0);
        row.CreateCell(2).SetCellValue(true);
        row.CreateCell(3).SetCellValue(new DateTime(2024, 1, 1));

        Assert.Equal("text", row.GetCell(0)!.StringCellValue);
        Assert.Equal(123.0, row.GetCell(1)!.NumericCellValue);
        Assert.True(row.GetCell(2)!.BooleanCellValue);
        Assert.True(row.GetCell(3)!.IsDateTimeCell);
    }
}
