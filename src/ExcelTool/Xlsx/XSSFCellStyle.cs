using ExcelTool.Interfaces;

namespace ExcelTool.Xlsx;

public class XSSFCellStyle : ICellStyle
{
    public HorizontalAlignment Alignment { get; set; } = HorizontalAlignment.General;
    public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Bottom;
    public bool WrapText { get; set; }

    public IFont? Font { get; set; }

    public string? FillForegroundColor { get; set; }
    public string? FillBackgroundColor { get; set; }
    public FillPattern FillPattern { get; set; } = FillPattern.NoFill;

    public BorderStyle BorderTop { get; set; } = BorderStyle.None;
    public BorderStyle BorderBottom { get; set; } = BorderStyle.None;
    public BorderStyle BorderLeft { get; set; } = BorderStyle.None;
    public BorderStyle BorderRight { get; set; } = BorderStyle.None;

    public string? TopBorderColor { get; set; }
    public string? BottomBorderColor { get; set; }
    public string? LeftBorderColor { get; set; }
    public string? RightBorderColor { get; set; }

    public string? DataFormat { get; set; }

    internal int StyleIndex { get; set; }
}
