using AutreMachine.ExcelTools.Interfaces;

namespace AutreMachine.ExcelTools.Xlsx;

public class XSSFWorkbook : IWorkbook
{
    private readonly List<XSSFSheet> _sheets = new();
    private readonly List<XSSFCellStyle> _styles = new();
    private readonly List<XSSFFont> _fonts = new();
    private int _sheetCounter = 1;
    private bool _disposed;

    public XSSFWorkbook()
    {
        _styles.Add(new XSSFCellStyle { StyleIndex = 0 });
        _fonts.Add(new XSSFFont { FontIndex = 0 });
    }

    /// <summary>Opens an existing XLSX workbook from a stream (NPOI-compatible constructor).</summary>
    public XSSFWorkbook(Stream stream) : this()
    {
        OpenXmlReader.Load(this, stream);
    }

    /// <summary>Opens an existing XLSX workbook from a file path (NPOI-compatible constructor).</summary>
    public XSSFWorkbook(string filePath) : this()
    {
        using var stream = File.OpenRead(filePath);
        OpenXmlReader.Load(this, stream);
    }

    /// <summary>Static factory — equivalent to <c>new XSSFWorkbook(stream)</c>.</summary>
    public static XSSFWorkbook Open(Stream stream) => new(stream);

    /// <summary>Static factory — equivalent to <c>new XSSFWorkbook(filePath)</c>.</summary>
    public static XSSFWorkbook Open(string filePath) => new(filePath);

    public int NumberOfSheets => _sheets.Count;

    public ISheet CreateSheet()
    {
        return CreateSheet($"Sheet{_sheetCounter++}");
    }

    public ISheet CreateSheet(string name)
    {
        if (_sheets.Any(s => s.SheetName.Equals(name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException($"A sheet named '{name}' already exists.");

        var sheet = new XSSFSheet(name);
        _sheets.Add(sheet);
        return sheet;
    }

    public ISheet GetSheetAt(int index)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 0);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, _sheets.Count);
        return _sheets[index];
    }

    public ISheet? GetSheet(string name) =>
        _sheets.FirstOrDefault(s => s.SheetName.Equals(name, StringComparison.OrdinalIgnoreCase));

    public int GetSheetIndex(string name)
    {
        for (int i = 0; i < _sheets.Count; i++)
            if (_sheets[i].SheetName.Equals(name, StringComparison.OrdinalIgnoreCase))
                return i;
        return -1;
    }

    public void RemoveSheetAt(int index)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 0);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, _sheets.Count);
        _sheets.RemoveAt(index);
    }

    public ICellStyle CreateCellStyle()
    {
        var style = new XSSFCellStyle { StyleIndex = _styles.Count };
        _styles.Add(style);
        return style;
    }

    public IFont CreateFont()
    {
        var font = new XSSFFont { FontIndex = _fonts.Count };
        _fonts.Add(font);
        return font;
    }

    public void Write(Stream stream) => OpenXmlSerializer.Serialize(this, stream);

    public void Write(string filePath)
    {
        using var stream = File.Create(filePath);
        Write(stream);
    }

    public void Close() => Dispose();

    public void Dispose()
    {
        if (!_disposed)
            _disposed = true;
        GC.SuppressFinalize(this);
    }

    internal IReadOnlyList<XSSFSheet> GetAllSheets() => _sheets;
    internal IReadOnlyList<XSSFCellStyle> GetAllStyles() => _styles;
    internal IReadOnlyList<XSSFFont> GetAllFonts() => _fonts;
}
