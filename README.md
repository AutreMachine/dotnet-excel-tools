# dotnet-excel-tools

It is a simple library to read and create Excel files, coded because NPOI was using a compromised DLL (Cryptography) that could not pass building chain.
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
