using ExcelTool.Interfaces;

namespace ExcelTool.Xlsx;

public class XSSFRow : IRow
{
    private readonly SortedDictionary<int, XSSFCell> _cells = new();

    internal XSSFRow(int rowNum)
    {
        RowNum = rowNum;
    }

    public int RowNum { get; }
    public double Height { get; set; }
    public short FirstCellNum => _cells.Count == 0 ? (short)-1 : (short)_cells.Keys.First();
    public short LastCellNum => _cells.Count == 0 ? (short)-1 : (short)(_cells.Keys.Last() + 1);

    public ICell CreateCell(int column)
    {
        var cell = new XSSFCell(RowNum, column);
        _cells[column] = cell;
        return cell;
    }

    public ICell CreateCell(int column, CellType type)
    {
        var cell = new XSSFCell(RowNum, column, type);
        _cells[column] = cell;
        return cell;
    }

    public ICell? GetCell(int column) =>
        _cells.TryGetValue(column, out var cell) ? cell : null;

    public void RemoveCell(ICell cell) =>
        _cells.Remove(cell.ColumnIndex);

    internal IEnumerable<XSSFCell> GetAllCells() => _cells.Values;
}
