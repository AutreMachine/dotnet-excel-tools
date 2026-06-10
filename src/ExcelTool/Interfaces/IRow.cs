namespace ExcelTool.Interfaces;

public interface IRow
{
    int RowNum { get; }
    short FirstCellNum { get; }
    short LastCellNum { get; }
    double Height { get; set; }

    ICell CreateCell(int column);
    ICell CreateCell(int column, CellType type);
    ICell? GetCell(int column);
    void RemoveCell(ICell cell);
}
