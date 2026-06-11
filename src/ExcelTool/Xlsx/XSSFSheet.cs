using AutreMachine.ExcelTools.Interfaces;

namespace AutreMachine.ExcelTools.Xlsx;

public class XSSFSheet : ISheet
{
    private readonly SortedDictionary<int, XSSFRow> _rows = new();
    private readonly Dictionary<int, int> _columnWidths = new();

    internal XSSFSheet(string name)
    {
        SheetName = name;
    }

    public string SheetName { get; }
    public int FirstRowNum => _rows.Count == 0 ? 0 : _rows.Keys.First();
    public int LastRowNum => _rows.Count == 0 ? 0 : _rows.Keys.Last();

    public IRow CreateRow(int rownum)
    {
        var row = new XSSFRow(rownum);
        _rows[rownum] = row;
        return row;
    }

    public IRow? GetRow(int rownum) =>
        _rows.TryGetValue(rownum, out var row) ? row : null;

    public void RemoveRow(IRow row) =>
        _rows.Remove(row.RowNum);

    public void SetColumnWidth(int columnIndex, int width) =>
        _columnWidths[columnIndex] = width;

    public void AutoSizeColumn(int columnIndex)
    {
        int maxLen = 10;
        foreach (var row in _rows.Values)
        {
            if (row.GetCell(columnIndex) is XSSFCell cell && cell.CellType == CellType.String)
                maxLen = Math.Max(maxLen, cell.StringCellValue.Length + 2);
        }
        _columnWidths[columnIndex] = Math.Min(maxLen * 256, 255 * 256);
    }

    internal IEnumerable<XSSFRow> GetAllRows() => _rows.Values;
    internal IReadOnlyDictionary<int, int> ColumnWidths => _columnWidths;
}
