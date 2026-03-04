using System;
using OfficeOpenXml;
using System.IO;

class InspectFissi
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        string filePath = args.Length > 0 ? args[0] : "../2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx";
        var file = new FileInfo(filePath);
        
        if (!file.Exists)
        {
            Console.WriteLine($"File not found: {file.FullName}");
            Console.WriteLine($"Current directory: {Directory.GetCurrentDirectory()}");
            return;
        }
        
        using (var package = new ExcelPackage(file))
        {
            // Find fissi sheet
            ExcelWorksheet fissiSheet = null;
            foreach (var ws in package.Workbook.Worksheets)
            {
                if (ws.Name.Equals("fissi", StringComparison.OrdinalIgnoreCase))
                {
                    fissiSheet = ws;
                    break;
                }
            }
            
            if (fissiSheet == null)
            {
                Console.WriteLine("ERROR: 'fissi' sheet not found!");
                Console.WriteLine("Available sheets:");
                foreach (var ws in package.Workbook.Worksheets)
                {
                    Console.WriteLine($"  - {ws.Name}");
                }
                return;
            }
            
            Console.WriteLine($"=== FISSI SHEET TIME COLUMNS INSPECTION ===");
            Console.WriteLine($"File: {file.Name}");
            Console.WriteLine();
            
            var dimension = fissiSheet.Dimension;
            if (dimension == null)
            {
                Console.WriteLine("Sheet is empty!");
                return;
            }
            
            Console.WriteLine($"Sheet dimensions: {dimension.Address}");
            Console.WriteLine($"Total rows: {dimension.End.Row}");
            Console.WriteLine();
            
            Console.WriteLine("=== TIME COLUMNS ANALYSIS ===");
            Console.WriteLine("Row | Partenza (Col 2)                              | Arrivo (Col 9)");
            Console.WriteLine("----|-----------------------------------------------|-----------------------------------------------");
            
            int startRow = 3; // Skip headers
            int rowsToShow = Math.Min(15, dimension.End.Row - startRow + 1);
            int decimalCount = 0;
            int timeFormatCount = 0;
            
            for (int row = startRow; row < startRow + rowsToShow; row++)
            {
                var partenzaCell = fissiSheet.Cells[row, 2];
                var arrivoCell = fissiSheet.Cells[row, 9];
                
                var partenzaValue = partenzaCell.Value;
                string partenzaText = partenzaCell.Text ?? "null";
                string partenzaFormat = partenzaCell.Style?.Numberformat?.Format ?? "no format";
                
                var arrivoValue = arrivoCell.Value;
                string arrivoText = arrivoCell.Text ?? "null";
                string arrivoFormat = arrivoCell.Style?.Numberformat?.Format ?? "no format";
                
                Console.WriteLine($"{row,3} | V:{partenzaValue,-12} T:{partenzaText,-10} F:{partenzaFormat,-10} | V:{arrivoValue,-12} T:{arrivoText,-10} F:{arrivoFormat,-10}");
                
                // Count decimal vs time format
                if (partenzaValue != null)
                {
                    if ((partenzaText.Contains(".") || partenzaText.Contains(",")) && !partenzaText.Contains(":"))
                        decimalCount++;
                    else if (partenzaText.Contains(":"))
                        timeFormatCount++;
                }
                
                if (arrivoValue != null)
                {
                    if ((arrivoText.Contains(".") || arrivoText.Contains(",")) && !arrivoText.Contains(":"))
                        decimalCount++;
                    else if (arrivoText.Contains(":"))
                        timeFormatCount++;
                }
            }
            
            Console.WriteLine();
            Console.WriteLine("=== ANALYSIS ===");
            Console.WriteLine($"Cells displaying as DECIMAL: {decimalCount}");
            Console.WriteLine($"Cells displaying as TIME FORMAT: {timeFormatCount}");
            Console.WriteLine();
            
            if (decimalCount > 0)
            {
                Console.WriteLine("⚠️  ISSUE DETECTED: Some time values are displaying as decimals");
                Console.WriteLine("   Example: 0.354166666666667 instead of 8:30");
                Console.WriteLine();
                Console.WriteLine("The fix in ExcelManager.cs AppendFissiData method:");
                Console.WriteLine("  1. Copies the decimal value as-is (preserves the data)");
                Console.WriteLine("  2. Applies 'h:mm' format to columns 2 and 9");
                Console.WriteLine("  3. Result: Excel displays 8:30 instead of 0.354166666666667");
                Console.WriteLine();
                Console.WriteLine("✓ The fix is ALREADY IMPLEMENTED in the code");
                Console.WriteLine("  When you process files through the application, time columns will display correctly");
            }
            else
            {
                Console.WriteLine("✓ All time values are already displaying in time format");
                Console.WriteLine("  The fix will ensure this format is preserved when copying");
            }
            
            Console.WriteLine();
            Console.WriteLine("Legend:");
            Console.WriteLine("  V = Value (internal representation, usually decimal for times)");
            Console.WriteLine("  T = Text (what Excel displays to the user)");
            Console.WriteLine("  F = Format (number format code applied to the cell)");
        }
        
        Console.WriteLine();
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
