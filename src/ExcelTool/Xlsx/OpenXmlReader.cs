using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OxCellFormat = DocumentFormat.OpenXml.Spreadsheet.CellFormat;

namespace ExcelTool.Xlsx;

internal static class OpenXmlReader
{
    internal static void Load(XSSFWorkbook workbook, Stream stream)
    {
        using var doc = SpreadsheetDocument.Open(stream, isEditable: false);
        var wbPart = doc.WorkbookPart
            ?? throw new InvalidDataException("The stream does not contain a valid Excel workbook.");

        var sharedStrings = ReadSharedStrings(wbPart);
        var dateStyleIndices = BuildDateStyleIndices(wbPart);

        foreach (Sheet sheet in wbPart.Workbook.Sheets!.Elements<Sheet>())
        {
            if (sheet.Id?.Value is null) continue;

            string sheetName = sheet.Name?.Value ?? $"Sheet{workbook.NumberOfSheets + 1}";
            var xssfSheet = (XSSFSheet)workbook.CreateSheet(sheetName);

            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id.Value);
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData is null) continue;

            foreach (Row row in sheetData.Elements<Row>())
            {
                if (row.RowIndex is null) continue;
                int rowNum = (int)(row.RowIndex.Value - 1);
                var xssfRow = (XSSFRow)xssfSheet.CreateRow(rowNum);

                if (row.CustomHeight?.Value == true && row.Height is not null)
                    xssfRow.Height = row.Height.Value;

                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference?.Value is null) continue;
                    int colIdx = ParseColumnIndex(cell.CellReference.Value);
                    var xssfCell = (XSSFCell)xssfRow.CreateCell(colIdx);

                    ReadCellValue(cell, xssfCell, sharedStrings, dateStyleIndices);
                }
            }
        }
    }

    // -------------------------------------------------------------------------
    // Cell value resolution
    // -------------------------------------------------------------------------

    private static void ReadCellValue(
        Cell cell, XSSFCell target,
        List<string> sharedStrings,
        HashSet<uint> dateStyleIndices)
    {
        var dataType = cell.DataType?.Value;
        if (dataType == CellValues.SharedString)
        {
            if (int.TryParse(cell.CellValue?.Text, out int ssIdx) && ssIdx < sharedStrings.Count)
                target.SetCellValue(sharedStrings[ssIdx]);
            return;
        }
        if (dataType == CellValues.Boolean)
        {
            target.SetCellValue(cell.CellValue?.Text == "1");
            return;
        }
        if (dataType == CellValues.InlineString)
        {
            target.SetCellValue(cell.GetFirstChild<InlineString>()?.Text?.Text ?? cell.InnerText);
            return;
        }

        // Numeric (or date, which is stored as a number in Excel)
        if (cell.CellValue?.Text is string raw &&
            double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
        {
            uint styleIdx = cell.StyleIndex?.Value ?? 0;
            if (dateStyleIndices.Contains(styleIdx))
                target.SetCellValue(DateTime.FromOADate(d));
            else
                target.SetCellValue(d);
        }
    }

    // -------------------------------------------------------------------------
    // Shared strings
    // -------------------------------------------------------------------------

    private static List<string> ReadSharedStrings(WorkbookPart wbPart)
    {
        var result = new List<string>();
        if (wbPart.SharedStringTablePart is null) return result;
        foreach (SharedStringItem item in
            wbPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
            result.Add(item.InnerText);
        return result;
    }

    // -------------------------------------------------------------------------
    // Date style detection
    // -------------------------------------------------------------------------

    private static HashSet<uint> BuildDateStyleIndices(WorkbookPart wbPart)
    {
        var result = new HashSet<uint>();
        var stylePart = wbPart.WorkbookStylesPart;
        if (stylePart?.Stylesheet is null) return result;

        // Custom numFmtId → format string
        var customFmts = new Dictionary<uint, string>();
        if (stylePart.Stylesheet.NumberingFormats is not null)
        {
            foreach (NumberingFormat nf in
                stylePart.Stylesheet.NumberingFormats.Elements<NumberingFormat>())
            {
                if (nf.NumberFormatId?.Value is not null && nf.FormatCode?.Value is not null)
                    customFmts[nf.NumberFormatId.Value] = nf.FormatCode.Value;
            }
        }

        if (stylePart.Stylesheet.CellFormats is null) return result;

        uint xfIdx = 0;
        foreach (OxCellFormat cf in stylePart.Stylesheet.CellFormats.Elements<OxCellFormat>())
        {
            uint numFmtId = cf.NumberFormatId?.Value ?? 0;
            bool isDate = IsBuiltInDateFmtId(numFmtId);
            if (!isDate && customFmts.TryGetValue(numFmtId, out string? fmt))
                isDate = LooksLikeDateFormat(fmt);
            if (isDate)
                result.Add(xfIdx);
            xfIdx++;
        }

        return result;
    }

    // Excel built-in number format IDs that represent dates / times
    private static bool IsBuiltInDateFmtId(uint id) =>
        (id >= 14 && id <= 17) || id == 22 ||
        (id >= 27 && id <= 36) || (id >= 45 && id <= 47);

    // Heuristic for custom format strings: presence of 'y' or 'd' tokens
    private static bool LooksLikeDateFormat(string format)
    {
        string stripped = Regex.Replace(format, "\"[^\"]*\"", "");
        return (stripped.Contains('y') || stripped.Contains('d')) &&
               !stripped.StartsWith('#') && !stripped.StartsWith('0');
    }

    // -------------------------------------------------------------------------
    // Cell reference → column index  ("A" → 0, "Z" → 25, "AA" → 26, …)
    // -------------------------------------------------------------------------

    private static int ParseColumnIndex(string cellRef)
    {
        int col = 0;
        foreach (char c in cellRef)
        {
            if (!char.IsLetter(c)) break;
            col = col * 26 + (c - 'A' + 1);
        }
        return col - 1;
    }
}
