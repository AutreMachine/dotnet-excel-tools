namespace ExcelTool.Interfaces;

public interface IWorkbook : IDisposable
{
    int NumberOfSheets { get; }

    ISheet CreateSheet();
    ISheet CreateSheet(string name);
    ISheet GetSheetAt(int index);
    ISheet? GetSheet(string name);
    int GetSheetIndex(string name);
    void RemoveSheetAt(int index);

    ICellStyle CreateCellStyle();
    IFont CreateFont();

    void Write(Stream stream);
    void Write(string filePath);
    void Close();
}
