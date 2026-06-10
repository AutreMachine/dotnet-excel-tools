using ExcelTool.Interfaces;

namespace ExcelTool.Xlsx;

public class XSSFCell : ICell
{
    private object? _value;

    internal XSSFCell(int rowIndex, int columnIndex, CellType cellType = CellType.Blank)
    {
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        CellType = cellType;
    }

    public int RowIndex { get; }
    public int ColumnIndex { get; }
    public CellType CellType { get; private set; }
    public ICellStyle? CellStyle { get; set; }

    public void SetCellValue(string value)
    {
        _value = value;
        CellType = CellType.String;
    }

    public void SetCellValue(double value)
    {
        _value = value;
        CellType = CellType.Numeric;
    }

    public void SetCellValue(bool value)
    {
        _value = value;
        CellType = CellType.Boolean;
    }

    public void SetCellValue(DateTime value)
    {
        _value = value;
        CellType = CellType.Numeric;
    }

    public void SetCellType(CellType cellType)
    {
        CellType = cellType;
        if (cellType == CellType.Blank)
            _value = null;
    }

    public void SetBlank()
    {
        _value = null;
        CellType = CellType.Blank;
    }

    public string StringCellValue =>
        CellType == CellType.String && _value is string s
            ? s
            : throw new InvalidOperationException($"Cell ({RowIndex},{ColumnIndex}) is not a string cell.");

    public double NumericCellValue =>
        CellType == CellType.Numeric && _value is not DateTime
            ? Convert.ToDouble(_value)
            : throw new InvalidOperationException($"Cell ({RowIndex},{ColumnIndex}) is not a numeric cell.");

    public bool BooleanCellValue =>
        CellType == CellType.Boolean && _value is bool b
            ? b
            : throw new InvalidOperationException($"Cell ({RowIndex},{ColumnIndex}) is not a boolean cell.");

    public DateTime DateCellValue =>
        _value is DateTime dt
            ? dt
            : throw new InvalidOperationException($"Cell ({RowIndex},{ColumnIndex}) does not contain a DateTime value.");

    public bool IsDateTimeCell => _value is DateTime;

    internal object? RawValue => _value;
    internal bool IsDateTime => _value is DateTime;
}
