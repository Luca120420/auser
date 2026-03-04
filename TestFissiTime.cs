using System;
using OfficeOpenXml;

class TestFissiTime
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        using (var package = new ExcelPackage(new System.IO.FileInfo("2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx")))
        {
            var fissiSheet = package.Workbook.Worksheets["fissi"];
            
            if (fissiSheet == null)
            {
                Console.WriteLine("fissi sheet not found");
                return;
            }
            
            Console.WriteLine("Inspecting fissi sheet:");
            Console.WriteLine("Row 3, Column 2 (Partenza):");
            var cell = fissiSheet.Cells[3, 2];
            Console.WriteLine($"  Value: {cell.Value}");
            Console.WriteLine($"  Value Type: {cell.Value?.GetType().Name}");
            Console.WriteLine($"  Text: {cell.Text}");
            Console.WriteLine($"  Number Format: '{cell.Style.Numberformat.Format}'");
            Console.WriteLine($"  Number Format ID: {cell.Style.Numberformat.NumFmtID}");
            
            Console.WriteLine("\nRow 3, Column 9 (Arrivo):");
            var cell2 = fissiSheet.Cells[3, 9];
            Console.WriteLine($"  Value: {cell2.Value}");
            Console.WriteLine($"  Value Type: {cell2.Value?.GetType().Name}");
            Console.WriteLine($"  Text: {cell2.Text}");
            Console.WriteLine($"  Number Format: '{cell2.Style.Numberformat.Format}'");
            Console.WriteLine($"  Number Format ID: {cell2.Style.Numberformat.NumFmtID}");
        }
    }
}
