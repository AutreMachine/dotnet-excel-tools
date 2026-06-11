# dotnet-excel-tools

It is a simple dotnet library to read and create Excel files, coded because NPOI was using a compromised DLL (Cryptography) that could not pass building chain.
It has been lazily 99% coded with Claude ! I really needed something quick - and not too dirty... so no credit for me.
Feel free to use ! There is only one library reference, it should be reliable in time.

Usage is close to NPOI.

* Creating a workbook :

```
// ─────────────────────────────────────────────────────────────────────────────
// Empty Excel file (one default sheet so Excel can open it)
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n[1] Creating an empty Excel file...");
using (IWorkbook wb = new XSSFWorkbook())
{
    wb.CreateSheet("Sheet1");
    string path = Path.Combine(outputDir, "test1_empty.xlsx");
    wb.Write(path);
    Console.WriteLine($"    -> {path}");
}
```

* Managing multiple worksheets :
```
// ─────────────────────────────────────────────────────────────────────────────
// Multiple worksheets in the same document
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n[2] Creating a file with multiple worksheets...");
using (IWorkbook wb = new XSSFWorkbook())
{
    string[] sheetNames = { "Overview", "Data", "Charts", "Config" };
    foreach (string name in sheetNames)
    {
        ISheet sheet = wb.CreateSheet(name);
        sheet.CreateRow(0).CreateCell(0).SetCellValue($"Sheet: {name}");
    }

    Console.WriteLine($"    Sheets created: {wb.NumberOfSheets}");
    for (int i = 0; i < wb.NumberOfSheets; i++)
        Console.WriteLine($"      [{i}] {wb.GetSheetAt(i).SheetName}");

    string path = Path.Combine(outputDir, "test2_multisheets.xlsx");
    wb.Write(path);
    Console.WriteLine($"    -> {path}");
}
```

* Creating different types of cells :

```
// ─────────────────────────────────────────────────────────────────────────────
// Different cell value types
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n[3] Writing cells with different types...");
using (IWorkbook wb = new XSSFWorkbook())
{
    ISheet sheet = wb.CreateSheet("CellTypes");

    IRow header = sheet.CreateRow(0);
    header.CreateCell(0).SetCellValue("Type");
    header.CreateCell(1).SetCellValue("Value");
    header.CreateCell(2).SetCellValue("Notes");

    IRow r1 = sheet.CreateRow(1);
    r1.CreateCell(0).SetCellValue("String");
    r1.CreateCell(1).SetCellValue("Hello, ExcelTool!");
    r1.CreateCell(2).SetCellValue("Plain text");

    IRow r2 = sheet.CreateRow(2);
    r2.CreateCell(0).SetCellValue("Integer");
    r2.CreateCell(1).SetCellValue(42.0);
    r2.CreateCell(2).SetCellValue("Whole number");

    IRow r3 = sheet.CreateRow(3);
    r3.CreateCell(0).SetCellValue("Decimal");
    r3.CreateCell(1).SetCellValue(3.14159265358979);
    r3.CreateCell(2).SetCellValue("Pi");

    IRow r4 = sheet.CreateRow(4);
    r4.CreateCell(0).SetCellValue("Negative");
    r4.CreateCell(1).SetCellValue(-9876.54);
    r4.CreateCell(2).SetCellValue("Negative double");

    IRow r5 = sheet.CreateRow(5);
    r5.CreateCell(0).SetCellValue("Boolean (true)");
    r5.CreateCell(1).SetCellValue(true);
    r5.CreateCell(2).SetCellValue("Logical TRUE");

    IRow r6 = sheet.CreateRow(6);
    r6.CreateCell(0).SetCellValue("Boolean (false)");
    r6.CreateCell(1).SetCellValue(false);
    r6.CreateCell(2).SetCellValue("Logical FALSE");

    // DateTime without explicit style — uses the built-in yyyy-mm-dd format
    IRow r7 = sheet.CreateRow(7);
    r7.CreateCell(0).SetCellValue("DateTime (auto-style)");
    r7.CreateCell(1).SetCellValue(new DateTime(2024, 6, 15, 14, 30, 0));
    r7.CreateCell(2).SetCellValue("Auto-formatted as yyyy-mm-dd");

    // DateTime with an explicit custom date style
    ICellStyle dateStyle = wb.CreateCellStyle();
    dateStyle.DataFormat = "dd/mm/yyyy hh:mm";
    IRow r8 = sheet.CreateRow(8);
    r8.CreateCell(0).SetCellValue("DateTime (custom style)");
    ICell dateCell = r8.CreateCell(1);
    dateCell.SetCellValue(new DateTime(2024, 12, 31, 23, 59, 59));
    dateCell.CellStyle = dateStyle;
    r8.CreateCell(2).SetCellValue("dd/mm/yyyy hh:mm");

    // Blank cell
    IRow r9 = sheet.CreateRow(9);
    r9.CreateCell(0).SetCellValue("Blank");
    r9.CreateCell(1, CellType.Blank);
    r9.CreateCell(2).SetCellValue("Intentionally empty");

    sheet.SetColumnWidth(0, 25 * 256);
    sheet.SetColumnWidth(1, 22 * 256);
    sheet.SetColumnWidth(2, 30 * 256);

    string path = Path.Combine(outputDir, "test3_celltypes.xlsx");
    wb.Write(path);
    Console.WriteLine($"    -> {path}");
}
```

* Styling cells :

```
// ─────────────────────────────────────────────────────────────────────────────
// Cell styles — fonts, fills, borders, alignment, number formats
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n[4] Writing styled cells...");
using (IWorkbook wb = new XSSFWorkbook())
{
    ISheet sheet = wb.CreateSheet("Styled");

    // Bold white-on-blue header
    IFont headerFont = wb.CreateFont();
    headerFont.IsBold = true;
    headerFont.FontHeightInPoints = 13;
    headerFont.Color = "FFFFFFFF";

    ICellStyle headerStyle = wb.CreateCellStyle();
    headerStyle.Font = headerFont;
    headerStyle.Alignment = HorizontalAlignment.Center;
    headerStyle.VerticalAlignment = VerticalAlignment.Center;
    headerStyle.FillForegroundColor = "FF2E75B6";
    headerStyle.FillPattern = FillPattern.SolidForeground;
    headerStyle.BorderBottom = BorderStyle.Medium;

    string[] headers = { "Name", "Score", "Grade", "Date", "Active" };
    IRow headerRow = sheet.CreateRow(0);
    headerRow.Height = 22;
    for (int c = 0; c < headers.Length; c++)
    {
        ICell hc = headerRow.CreateCell(c);
        hc.SetCellValue(headers[c]);
        hc.CellStyle = headerStyle;
    }

    // Thin-border base style
    ICellStyle dataStyle = wb.CreateCellStyle();
    dataStyle.BorderTop = BorderStyle.Thin;
    dataStyle.BorderBottom = BorderStyle.Thin;
    dataStyle.BorderLeft = BorderStyle.Thin;
    dataStyle.BorderRight = BorderStyle.Thin;

    // Numeric format
    ICellStyle numStyle = wb.CreateCellStyle();
    numStyle.DataFormat = "#,##0.00";
    numStyle.Alignment = HorizontalAlignment.Right;
    numStyle.BorderTop = BorderStyle.Thin;
    numStyle.BorderBottom = BorderStyle.Thin;
    numStyle.BorderLeft = BorderStyle.Thin;
    numStyle.BorderRight = BorderStyle.Thin;

    // Date style
    ICellStyle dateStyle2 = wb.CreateCellStyle();
    dateStyle2.DataFormat = "dd-mmm-yyyy";
    dateStyle2.Alignment = HorizontalAlignment.Center;
    dateStyle2.BorderTop = BorderStyle.Thin;
    dateStyle2.BorderBottom = BorderStyle.Thin;
    dateStyle2.BorderLeft = BorderStyle.Thin;
    dateStyle2.BorderRight = BorderStyle.Thin;

    // Italic red for low scores
    IFont redFont = wb.CreateFont();
    redFont.IsItalic = true;
    redFont.Color = "FFCC0000";

    ICellStyle redNumStyle = wb.CreateCellStyle();
    redNumStyle.Font = redFont;
    redNumStyle.DataFormat = "#,##0.00";
    redNumStyle.Alignment = HorizontalAlignment.Right;
    redNumStyle.BorderTop = BorderStyle.Thin;
    redNumStyle.BorderBottom = BorderStyle.Thin;
    redNumStyle.BorderLeft = BorderStyle.Thin;
    redNumStyle.BorderRight = BorderStyle.Thin;

    // Alternating row background styles
    ICellStyle altStyle = wb.CreateCellStyle();
    altStyle.FillForegroundColor = "FFE9EFF7";
    altStyle.FillPattern = FillPattern.SolidForeground;
    altStyle.BorderTop = BorderStyle.Thin;
    altStyle.BorderBottom = BorderStyle.Thin;
    altStyle.BorderLeft = BorderStyle.Thin;
    altStyle.BorderRight = BorderStyle.Thin;

    ICellStyle altNumStyle = wb.CreateCellStyle();
    altNumStyle.FillForegroundColor = "FFE9EFF7";
    altNumStyle.FillPattern = FillPattern.SolidForeground;
    altNumStyle.DataFormat = "#,##0.00";
    altNumStyle.Alignment = HorizontalAlignment.Right;
    altNumStyle.BorderTop = BorderStyle.Thin;
    altNumStyle.BorderBottom = BorderStyle.Thin;
    altNumStyle.BorderLeft = BorderStyle.Thin;
    altNumStyle.BorderRight = BorderStyle.Thin;

    ICellStyle altDateStyle = wb.CreateCellStyle();
    altDateStyle.FillForegroundColor = "FFE9EFF7";
    altDateStyle.FillPattern = FillPattern.SolidForeground;
    altDateStyle.DataFormat = "dd-mmm-yyyy";
    altDateStyle.Alignment = HorizontalAlignment.Center;
    altDateStyle.BorderTop = BorderStyle.Thin;
    altDateStyle.BorderBottom = BorderStyle.Thin;
    altDateStyle.BorderLeft = BorderStyle.Thin;
    altDateStyle.BorderRight = BorderStyle.Thin;

    var records = new (string name, double score, string grade, DateTime date, bool active)[]
    {
        ("Alice",  98.50, "A+", new DateTime(2024,  1, 10), true),
        ("Bob",    54.30, "D",  new DateTime(2024,  2, 14), false),
        ("Carol",  87.00, "B+", new DateTime(2024,  3, 22), true),
        ("David",  43.75, "F",  new DateTime(2024,  4,  5), false),
        ("Eve",    92.10, "A",  new DateTime(2024,  5, 30), true),
    };

    for (int i = 0; i < records.Length; i++)
    {
        var (name, score, grade, date, active) = records[i];
        bool alt = i % 2 == 1;
        IRow row = sheet.CreateRow(i + 1);

        ICellStyle textSt  = alt ? altStyle : dataStyle;
        ICellStyle scoreSt = score < 60 ? redNumStyle : (alt ? altNumStyle : numStyle);
        ICellStyle dateSt  = alt ? altDateStyle : dateStyle2;

        ICell nc = row.CreateCell(0); nc.SetCellValue(name);  nc.CellStyle = textSt;
        ICell sc = row.CreateCell(1); sc.SetCellValue(score); sc.CellStyle = scoreSt;
        ICell gc = row.CreateCell(2); gc.SetCellValue(grade); gc.CellStyle = textSt;
        ICell dc = row.CreateCell(3); dc.SetCellValue(date);  dc.CellStyle = dateSt;
        ICell ac = row.CreateCell(4); ac.SetCellValue(active);ac.CellStyle = textSt;
    }

    sheet.SetColumnWidth(0, 16 * 256);
    sheet.SetColumnWidth(1, 12 * 256);
    sheet.SetColumnWidth(2, 10 * 256);
    sheet.SetColumnWidth(3, 15 * 256);
    sheet.SetColumnWidth(4, 10 * 256);

    string path = Path.Combine(outputDir, "test4_styled.xlsx");
    wb.Write(path);
    Console.WriteLine($"    -> {path}");
}
```

* Loading a workbook :

```
// ─────────────────────────────────────────────────────────────────────────────
// Load an existing workbook and read its cell values back
// ─────────────────────────────────────────────────────────────────────────────
Console.WriteLine("\n[6] Loading an existing workbook and reading cells back...");
{
    string sourcePath = Path.Combine(outputDir, "test3_celltypes.xlsx");
    using IWorkbook wb = new XSSFWorkbook(sourcePath);      // NPOI-style constructor
    // Same result: using IWorkbook wb = XSSFWorkbook.Open(sourcePath);

    Console.WriteLine($"    Loaded '{sourcePath}'");
    Console.WriteLine($"    Sheets: {wb.NumberOfSheets}");

    ISheet sheet = wb.GetSheetAt(0);
    Console.WriteLine($"    Sheet name: {sheet.SheetName}");
    Console.WriteLine($"    Rows: {sheet.LastRowNum - sheet.FirstRowNum + 1}");

    // Read and print each data row (skip header row 0)
    for (int r = 1; r <= sheet.LastRowNum; r++)
    {
        IRow? row = sheet.GetRow(r);
        if (row is null) continue;

        ICell? typeCell  = row.GetCell(0);
        ICell? valueCell = row.GetCell(1);
        if (typeCell is null || valueCell is null) continue;

        string typeName = typeCell.StringCellValue;
        string valueStr = valueCell.CellType switch
        {
            CellType.String  => valueCell.StringCellValue,
            CellType.Numeric => valueCell.IsDateTimeCell
                                    ? valueCell.DateCellValue.ToString("yyyy-MM-dd HH:mm:ss")
                                    : valueCell.NumericCellValue.ToString("G"),
            CellType.Boolean => valueCell.BooleanCellValue.ToString(),
            CellType.Blank   => "(blank)",
            _                => $"({valueCell.CellType})"
        };

        Console.WriteLine($"      row {r}: [{typeName}] = {valueStr}");
    }
}
```
