using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OxBorder = DocumentFormat.OpenXml.Spreadsheet.Border;
using OxFont = DocumentFormat.OpenXml.Spreadsheet.Font;
using OxColor = DocumentFormat.OpenXml.Spreadsheet.Color;
using OxBorderStyle = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues;

namespace ExcelTool.Xlsx;

internal static class OpenXmlSerializer
{
    public static void Serialize(XSSFWorkbook workbook, Stream stream)
    {
        using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var sharedStrings = new List<string>();

        var (stylesheet, styleIndexMap, defaultDateStyleIdx) = BuildStylesheet(workbook);
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = stylesheet;
        stylesPart.Stylesheet.Save();

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        uint sheetId = 1;

        foreach (var sheet in workbook.GetAllSheets())
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var worksheet = new Worksheet();

            var cols = BuildColumns(sheet);
            if (cols is not null)
                worksheet.AppendChild(cols);

            var sheetData = new SheetData();
            foreach (var row in sheet.GetAllRows())
                sheetData.AppendChild(BuildRow(row, sharedStrings, styleIndexMap, defaultDateStyleIdx));

            worksheet.AppendChild(sheetData);
            worksheetPart.Worksheet = worksheet;
            worksheetPart.Worksheet.Save();

            sheets.AppendChild(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId++,
                Name = sheet.SheetName
            });
        }

        if (sharedStrings.Count > 0)
        {
            var sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
            var sst = new SharedStringTable();
            foreach (var s in sharedStrings)
                sst.AppendChild(new SharedStringItem(
                    new Text(s) { Space = SpaceProcessingModeValues.Preserve }));
            sstPart.SharedStringTable = sst;
            sstPart.SharedStringTable.Save();
        }

        workbookPart.Workbook.Save();
    }

    // -------------------------------------------------------------------------
    // Row / Cell builders
    // -------------------------------------------------------------------------

    private static Row BuildRow(
        XSSFRow row,
        List<string> sharedStrings,
        Dictionary<int, uint> styleIndexMap,
        uint defaultDateStyleIdx)
    {
        var openXmlRow = new Row { RowIndex = (uint)(row.RowNum + 1) };
        if (row.Height > 0)
        {
            openXmlRow.Height = row.Height;
            openXmlRow.CustomHeight = true;
        }

        foreach (var cell in row.GetAllCells())
        {
            if (cell.CellType == CellType.Blank && cell.CellStyle is null)
                continue;

            openXmlRow.AppendChild(BuildCell(cell, sharedStrings, styleIndexMap, defaultDateStyleIdx));
        }

        return openXmlRow;
    }

    private static Cell BuildCell(
        XSSFCell cell,
        List<string> sharedStrings,
        Dictionary<int, uint> styleIndexMap,
        uint defaultDateStyleIdx)
    {
        var openXmlCell = new Cell { CellReference = CellReference(cell.RowIndex, cell.ColumnIndex) };

        // Apply explicit style first
        if (cell.CellStyle is XSSFCellStyle s && styleIndexMap.TryGetValue(s.StyleIndex, out uint sIdx))
            openXmlCell.StyleIndex = sIdx;

        switch (cell.CellType)
        {
            case CellType.String:
            {
                string val = cell.StringCellValue;
                int idx = sharedStrings.IndexOf(val);
                if (idx < 0) { idx = sharedStrings.Count; sharedStrings.Add(val); }
                openXmlCell.DataType = CellValues.SharedString;
                openXmlCell.CellValue = new CellValue(idx.ToString());
                break;
            }
            case CellType.Numeric when cell.IsDateTime:
            {
                openXmlCell.CellValue = new CellValue(
                    cell.DateCellValue.ToOADate().ToString(CultureInfo.InvariantCulture));
                // Use the default date style only when no explicit style was set
                if (cell.CellStyle is null)
                    openXmlCell.StyleIndex = defaultDateStyleIdx;
                break;
            }
            case CellType.Numeric:
                openXmlCell.CellValue = new CellValue(
                    cell.NumericCellValue.ToString(CultureInfo.InvariantCulture));
                break;
            case CellType.Boolean:
                openXmlCell.DataType = CellValues.Boolean;
                openXmlCell.CellValue = new CellValue(cell.BooleanCellValue ? "1" : "0");
                break;
        }

        return openXmlCell;
    }

    // -------------------------------------------------------------------------
    // Stylesheet builder
    // -------------------------------------------------------------------------

    private static (Stylesheet stylesheet, Dictionary<int, uint> styleIndexMap, uint defaultDateStyleIdx)
        BuildStylesheet(XSSFWorkbook workbook)
    {
        var numFmts = new NumberingFormats();
        var fonts = new Fonts();
        var fills = new Fills();
        var borders = new Borders();
        var cellStyleXfs = new CellStyleFormats();
        var cellXfs = new CellFormats();

        uint nextNumFmtId = 164;
        var numFmtCache = new Dictionary<string, uint>();

        // ----- Required defaults -----
        fonts.AppendChild(DefaultFont());
        fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.None }));
        fills.AppendChild(new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
        borders.AppendChild(DefaultBorder());
        cellStyleXfs.AppendChild(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 });
        cellXfs.AppendChild(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 });

        // ----- Font index map (workbook font index → OpenXML index) -----
        var fontIdxMap = new Dictionary<int, uint> { [0] = 0 };
        foreach (var f in workbook.GetAllFonts().Skip(1))
        {
            uint fi = (uint)fonts.ChildElements.Count;
            fonts.AppendChild(BuildFont(f));
            fontIdxMap[f.FontIndex] = fi;
        }

        // ----- Default date format (always present so DateTime cells work without a style) -----
        const string defaultDateFmtStr = "yyyy\\-mm\\-dd";
        uint dateFmtId = nextNumFmtId++;
        numFmts.AppendChild(new NumberingFormat { NumberFormatId = dateFmtId, FormatCode = defaultDateFmtStr });
        numFmtCache[defaultDateFmtStr] = dateFmtId;

        uint defaultDateStyleIdx = (uint)cellXfs.ChildElements.Count;
        cellXfs.AppendChild(new CellFormat
        {
            NumberFormatId = dateFmtId, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0,
            ApplyNumberFormat = true
        });

        // ----- Per-style fill / border caches -----
        var fillCache = new Dictionary<string, uint>();
        var borderCache = new Dictionary<string, uint>();
        var styleIndexMap = new Dictionary<int, uint> { [0] = 0 };

        foreach (var style in workbook.GetAllStyles().Skip(1))
        {
            // Font
            uint fontId = 0;
            if (style.Font is XSSFFont styleFont && fontIdxMap.TryGetValue(styleFont.FontIndex, out uint fi2))
                fontId = fi2;

            // Fill
            uint fillId = 0;
            if (style.FillPattern != FillPattern.NoFill && style.FillForegroundColor is not null)
            {
                string fillKey = $"{style.FillPattern}|{style.FillForegroundColor}|{style.FillBackgroundColor}";
                if (!fillCache.TryGetValue(fillKey, out fillId))
                {
                    fillId = (uint)fills.ChildElements.Count;
                    fills.AppendChild(BuildFill(style));
                    fillCache[fillKey] = fillId;
                }
            }

            // Border
            uint borderId = 0;
            bool hasBorder = style.BorderTop != BorderStyle.None || style.BorderBottom != BorderStyle.None
                          || style.BorderLeft != BorderStyle.None || style.BorderRight != BorderStyle.None;
            if (hasBorder)
            {
                string borderKey =
                    $"{style.BorderLeft}|{style.BorderRight}|{style.BorderTop}|{style.BorderBottom}";
                if (!borderCache.TryGetValue(borderKey, out borderId))
                {
                    borderId = (uint)borders.ChildElements.Count;
                    borders.AppendChild(BuildBorder(style));
                    borderCache[borderKey] = borderId;
                }
            }

            // Number format
            uint numFmtId = 0;
            bool applyNumFmt = false;
            if (!string.IsNullOrEmpty(style.DataFormat))
            {
                if (!numFmtCache.TryGetValue(style.DataFormat, out numFmtId))
                {
                    numFmtId = nextNumFmtId++;
                    numFmts.AppendChild(new NumberingFormat
                    {
                        NumberFormatId = numFmtId,
                        FormatCode = style.DataFormat
                    });
                    numFmtCache[style.DataFormat] = numFmtId;
                }
                applyNumFmt = true;
            }

            // Alignment
            Alignment? alignment = null;
            bool hasAlignment = style.Alignment != HorizontalAlignment.General
                             || style.VerticalAlignment != VerticalAlignment.Bottom
                             || style.WrapText;
            if (hasAlignment)
            {
                alignment = new Alignment
                {
                    Horizontal = MapHAlign(style.Alignment),
                    Vertical = MapVAlign(style.VerticalAlignment)
                };
                if (style.WrapText) alignment.WrapText = true;
            }

            var fmt = new CellFormat
            {
                NumberFormatId = numFmtId,
                FontId = fontId,
                FillId = fillId,
                BorderId = borderId,
                FormatId = 0
            };
            if (fontId > 0 || style.Font is not null) fmt.ApplyFont = true;
            if (fillId > 0) fmt.ApplyFill = true;
            if (borderId > 0) fmt.ApplyBorder = true;
            if (applyNumFmt) fmt.ApplyNumberFormat = true;
            if (alignment is not null) { fmt.ApplyAlignment = true; fmt.Alignment = alignment; }

            uint xfIdx = (uint)cellXfs.ChildElements.Count;
            cellXfs.AppendChild(fmt);
            styleIndexMap[style.StyleIndex] = xfIdx;
        }

        fonts.Count = (uint)fonts.ChildElements.Count;
        fills.Count = (uint)fills.ChildElements.Count;
        borders.Count = (uint)borders.ChildElements.Count;
        cellStyleXfs.Count = (uint)cellStyleXfs.ChildElements.Count;
        cellXfs.Count = (uint)cellXfs.ChildElements.Count;
        numFmts.Count = (uint)numFmts.ChildElements.Count;

        var stylesheet = new Stylesheet();
        stylesheet.AppendChild(numFmts);
        stylesheet.AppendChild(fonts);
        stylesheet.AppendChild(fills);
        stylesheet.AppendChild(borders);
        stylesheet.AppendChild(cellStyleXfs);
        stylesheet.AppendChild(cellXfs);

        return (stylesheet, styleIndexMap, defaultDateStyleIdx);
    }

    // -------------------------------------------------------------------------
    // Element factory helpers
    // -------------------------------------------------------------------------

    private static OxFont DefaultFont() =>
        new OxFont(new FontSize { Val = 11 }, new FontName { Val = "Calibri" });

    private static OxFont BuildFont(XSSFFont f)
    {
        var font = new OxFont();
        font.AppendChild(new FontSize { Val = f.FontHeightInPoints });
        font.AppendChild(new FontName { Val = f.FontName });
        if (f.IsBold) font.AppendChild(new Bold());
        if (f.IsItalic) font.AppendChild(new Italic());
        if (f.IsStrikeout) font.AppendChild(new Strike());
        if (f.Underline != FontUnderlineType.None)
            font.AppendChild(new Underline { Val = MapUnderline(f.Underline) });
        if (f.Color != "FF000000")
            font.AppendChild(new OxColor { Rgb = f.Color });
        return font;
    }

    private static OxBorder DefaultBorder() =>
        new OxBorder(new LeftBorder(), new RightBorder(), new TopBorder(),
                     new BottomBorder(), new DiagonalBorder());

    private static OxBorder BuildBorder(XSSFCellStyle s)
    {
        var b = new OxBorder();
        b.AppendChild(MakeBorderPart<LeftBorder>(s.BorderLeft, s.LeftBorderColor));
        b.AppendChild(MakeBorderPart<RightBorder>(s.BorderRight, s.RightBorderColor));
        b.AppendChild(MakeBorderPart<TopBorder>(s.BorderTop, s.TopBorderColor));
        b.AppendChild(MakeBorderPart<BottomBorder>(s.BorderBottom, s.BottomBorderColor));
        b.AppendChild(new DiagonalBorder());
        return b;
    }

    private static T MakeBorderPart<T>(BorderStyle style, string? color)
        where T : BorderPropertiesType, new()
    {
        var part = new T();
        if (style != BorderStyle.None)
        {
            part.Style = MapBorderStyle(style);
            if (color is not null)
                part.AppendChild(new OxColor { Rgb = color });
        }
        return part;
    }

    private static Fill BuildFill(XSSFCellStyle s)
    {
        var pf = new PatternFill { PatternType = MapFillPattern(s.FillPattern) };
        if (s.FillForegroundColor is not null)
            pf.AppendChild(new ForegroundColor { Rgb = s.FillForegroundColor });
        if (s.FillBackgroundColor is not null)
            pf.AppendChild(new BackgroundColor { Rgb = s.FillBackgroundColor });
        return new Fill(pf);
    }

    private static Columns? BuildColumns(XSSFSheet sheet)
    {
        if (!sheet.ColumnWidths.Any()) return null;
        var cols = new Columns();
        foreach (var (colIdx, width) in sheet.ColumnWidths)
        {
            cols.AppendChild(new Column
            {
                Min = (uint)(colIdx + 1),
                Max = (uint)(colIdx + 1),
                Width = width / 256.0,
                CustomWidth = true
            });
        }
        return cols;
    }

    // -------------------------------------------------------------------------
    // Coordinate helpers
    // -------------------------------------------------------------------------

    private static string CellReference(int rowIndex, int colIndex) =>
        ColumnLetter(colIndex) + (rowIndex + 1);

    private static string ColumnLetter(int colIndex)
    {
        string result = "";
        colIndex++;
        while (colIndex > 0)
        {
            colIndex--;
            result = (char)('A' + colIndex % 26) + result;
            colIndex /= 26;
        }
        return result;
    }

    // -------------------------------------------------------------------------
    // Enum mappers
    // -------------------------------------------------------------------------

    private static HorizontalAlignmentValues MapHAlign(HorizontalAlignment a) => a switch
    {
        HorizontalAlignment.Left => HorizontalAlignmentValues.Left,
        HorizontalAlignment.Center => HorizontalAlignmentValues.Center,
        HorizontalAlignment.Right => HorizontalAlignmentValues.Right,
        HorizontalAlignment.Fill => HorizontalAlignmentValues.Fill,
        HorizontalAlignment.Justify => HorizontalAlignmentValues.Justify,
        HorizontalAlignment.CenterSelection => HorizontalAlignmentValues.CenterContinuous,
        HorizontalAlignment.Distributed => HorizontalAlignmentValues.Distributed,
        _ => HorizontalAlignmentValues.General
    };

    private static VerticalAlignmentValues MapVAlign(VerticalAlignment v) => v switch
    {
        VerticalAlignment.Top => VerticalAlignmentValues.Top,
        VerticalAlignment.Center => VerticalAlignmentValues.Center,
        VerticalAlignment.Justify => VerticalAlignmentValues.Justify,
        VerticalAlignment.Distributed => VerticalAlignmentValues.Distributed,
        _ => VerticalAlignmentValues.Bottom
    };

    private static OxBorderStyle MapBorderStyle(BorderStyle s) => s switch
    {
        BorderStyle.Thin => OxBorderStyle.Thin,
        BorderStyle.Medium => OxBorderStyle.Medium,
        BorderStyle.Dashed => OxBorderStyle.Dashed,
        BorderStyle.Dotted => OxBorderStyle.Dotted,
        BorderStyle.Thick => OxBorderStyle.Thick,
        BorderStyle.Double => OxBorderStyle.Double,
        BorderStyle.Hair => OxBorderStyle.Hair,
        BorderStyle.MediumDashed => OxBorderStyle.MediumDashed,
        BorderStyle.DashDot => OxBorderStyle.DashDot,
        BorderStyle.MediumDashDot => OxBorderStyle.MediumDashDot,
        BorderStyle.DashDotDot => OxBorderStyle.DashDotDot,
        BorderStyle.MediumDashDotDot => OxBorderStyle.MediumDashDotDot,
        BorderStyle.SlantedDashDot => OxBorderStyle.SlantDashDot,
        _ => OxBorderStyle.None
    };

    private static PatternValues MapFillPattern(FillPattern p) => p switch
    {
        FillPattern.SolidForeground => PatternValues.Solid,
        FillPattern.FineDots => PatternValues.Gray0625,
        FillPattern.SparseDots => PatternValues.Gray125,
        FillPattern.ThickHorizontalBands => PatternValues.MediumGray,
        FillPattern.ThickVerticalBands => PatternValues.DarkVertical,
        FillPattern.ThinHorizontalBands => PatternValues.LightHorizontal,
        FillPattern.ThinVerticalBands => PatternValues.LightVertical,
        FillPattern.ThinBackwardDiagonals => PatternValues.LightDown,
        FillPattern.ThinForwardDiagonals => PatternValues.LightUp,
        FillPattern.Squares => PatternValues.LightTrellis,
        FillPattern.Diamonds => PatternValues.LightGrid,
        _ => PatternValues.None
    };

    private static UnderlineValues MapUnderline(FontUnderlineType u) => u switch
    {
        FontUnderlineType.Double => UnderlineValues.Double,
        FontUnderlineType.SingleAccounting => UnderlineValues.SingleAccounting,
        FontUnderlineType.DoubleAccounting => UnderlineValues.DoubleAccounting,
        _ => UnderlineValues.Single
    };
}
