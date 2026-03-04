using System;
using System.IO;
using OfficeOpenXml;

namespace VerifyTimeFix;

class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Create a test workbook
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Test");
            
            // Simulate what the code does:
            // 1. Source has DateTime value
            DateTime sourceDateTime = new DateTime(1899, 12, 30, 8, 30, 0);
            
            // 2. Convert to TotalDays (what the current code does)
            double timeValue = sourceDateTime.TimeOfDay.TotalDays;
            
            // 3. Set value and format
            ws.Cells[1, 1].Value = timeValue;
            ws.Cells[1, 1].Style.Numberformat.Format = "h:mm";
            
            // Also test with direct DateTime
            ws.Cells[2, 1].Value = sourceDateTime;
            ws.Cells[2, 1].Style.Numberformat.Format = "h:mm";
            
            // Save
            package.SaveAs(new FileInfo("../TimeFormatTest.xlsx"));
        }

        // Read back and verify
        using (var package = new ExcelPackage(new FileInfo("../TimeFormatTest.xlsx")))
        {
            var ws = package.Workbook.Worksheets["Test"];
            
            Console.WriteLine("=== VERIFICATION ===");
            Console.WriteLine($"Cell A1 (TotalDays approach):");
            Console.WriteLine($"  Value: {ws.Cells[1, 1].Value}");
            Console.WriteLine($"  Text: {ws.Cells[1, 1].Text}");
            Console.WriteLine($"  Format: {ws.Cells[1, 1].Style.Numberformat.Format}");
            
            Console.WriteLine($"\nCell A2 (Direct DateTime):");
            Console.WriteLine($"  Value: {ws.Cells[2, 1].Value}");
            Console.WriteLine($"  Text: {ws.Cells[2, 1].Text}");
            Console.WriteLine($"  Format: {ws.Cells[2, 1].Style.Numberformat.Format}");
            
            Console.WriteLine("\n✓ Test file created: TimeFormatTest.xlsx");
            Console.WriteLine("Open it in Excel to verify the display");
        }
    }
}
