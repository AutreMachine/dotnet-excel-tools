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
