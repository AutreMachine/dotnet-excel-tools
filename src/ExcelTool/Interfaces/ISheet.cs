namespace ExcelTool.Interfaces;

public interface ISheet
{
    string SheetName { get; }
    int FirstRowNum { get; }
    int LastRowNum { get; }

    IRow CreateRow(int rownum);
    IRow? GetRow(int rownum);
    void RemoveRow(IRow row);

    /// <summary>Width in 1/256th of a character unit, matching NPOI convention.</summary>
    void SetColumnWidth(int columnIndex, int width);
    void AutoSizeColumn(int columnIndex);
}
