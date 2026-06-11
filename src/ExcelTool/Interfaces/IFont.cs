namespace AutreMachine.ExcelTools.Interfaces;

public interface IFont
{
    string FontName { get; set; }
    short FontHeightInPoints { get; set; }
    bool IsBold { get; set; }
    bool IsItalic { get; set; }
    bool IsStrikeout { get; set; }
    FontUnderlineType Underline { get; set; }
    /// <summary>ARGB hex color, e.g. "FF000000" for black.</summary>
    string Color { get; set; }
}
