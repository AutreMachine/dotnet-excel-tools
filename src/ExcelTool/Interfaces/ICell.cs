namespace AutreMachine.ExcelTools.Interfaces;

public interface ICell
{
    int RowIndex { get; }
    int ColumnIndex { get; }
    CellType CellType { get; }
    ICellStyle? CellStyle { get; set; }

    void SetCellValue(string value);
    void SetCellValue(double value);
    void SetCellValue(bool value);
    void SetCellValue(DateTime value);
    void SetCellType(CellType cellType);
    void SetBlank();

    string StringCellValue { get; }
    double NumericCellValue { get; }
    bool BooleanCellValue { get; }
    DateTime DateCellValue { get; }
    /// <summary>True when the cell holds a DateTime value (CellType is Numeric but stored as a date).</summary>
    bool IsDateTimeCell { get; }
}
