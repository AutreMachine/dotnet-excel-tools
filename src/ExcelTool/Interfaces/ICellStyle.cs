namespace ExcelTool.Interfaces;

public interface ICellStyle
{
    HorizontalAlignment Alignment { get; set; }
    VerticalAlignment VerticalAlignment { get; set; }
    bool WrapText { get; set; }

    IFont? Font { get; set; }

    /// <summary>ARGB hex color, e.g. "FF4472C4". Null means no fill.</summary>
    string? FillForegroundColor { get; set; }
    string? FillBackgroundColor { get; set; }
    FillPattern FillPattern { get; set; }

    BorderStyle BorderTop { get; set; }
    BorderStyle BorderBottom { get; set; }
    BorderStyle BorderLeft { get; set; }
    BorderStyle BorderRight { get; set; }

    /// <summary>ARGB hex color for each border edge. Null uses Excel default (black).</summary>
    string? TopBorderColor { get; set; }
    string? BottomBorderColor { get; set; }
    string? LeftBorderColor { get; set; }
    string? RightBorderColor { get; set; }

    /// <summary>Excel number format string, e.g. "yyyy-mm-dd" or "#,##0.00".</summary>
    string? DataFormat { get; set; }
}
