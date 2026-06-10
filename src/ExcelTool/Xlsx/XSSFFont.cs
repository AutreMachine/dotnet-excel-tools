using ExcelTool.Interfaces;

namespace ExcelTool.Xlsx;

public class XSSFFont : IFont
{
    public string FontName { get; set; } = "Calibri";
    public short FontHeightInPoints { get; set; } = 11;
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool IsStrikeout { get; set; }
    public FontUnderlineType Underline { get; set; } = FontUnderlineType.None;
    public string Color { get; set; } = "FF000000";

    internal int FontIndex { get; set; }
}
