using System;
using System.Linq;
using AuserExcelTransformer.Services;

class QuickTest
{
    static void Main()
    {
        Console.WriteLine("=== QUICK TEST ===\n");
        
        var csvParser = new CSVParser();
        var transformationRulesEngine = new TransformationRulesEngine();
        var dataTransformer = new DataTransformer(transformationRulesEngine);
        var excelManager = new ExcelManager();
        
        // Test 1: Parse CSV
        Console.WriteLine("Test 1: Parsing CSV...");
        var appointments = csvParser.ParseCSV("168514-Estrazione_1770193162042.csv");
        Console.WriteLine($"✓ Parsed {appointments.Count} appointments\n");
        
        // Test 2: Transform data
        Console.WriteLine("Test 2: Transforming data...");
        var result = dataTransformer.Transform(appointments);
        Console.WriteLine($"✓ Transformed to {result.Rows.Count} rows");
        Console.WriteLine($"✓ Yellow highlight rows: {result.YellowHighlightRows.Count}\n");
        
        // Test 3: Check first row mapping
        Console.WriteLine("Test 3: Checking CSV column mapping...");
        if (result.Rows.Count > 0)
        {
            var firstRow = result.Rows[0];
            Console.WriteLine($"  Data: {firstRow.DataServizio}");
            Console.WriteLine($"  Ora Inizio Servizio: '{firstRow.OraInizioServizio}' (should be EMPTY)");
            Console.WriteLine($"  Assistito: {firstRow.Assistito}");
            Console.WriteLine($"  Destinazione: {firstRow.Destinazione}");
            Console.WriteLine($"  Note: {firstRow.Note}");
            Console.WriteLine($"  Auto: {firstRow.Auto}");
            Console.WriteLine($"  Volontario: {firstRow.Volontario}");
            Console.WriteLine($"  Arrivo: {firstRow.Arrivo}");
            Console.WriteLine($"  Indirizzo Partenza: {firstRow.IndirizzoPartenza}");
            Console.WriteLine($"  Comune Partenza: {firstRow.ComunePartenza}");
            Console.WriteLine($"  Indirizzo Gasnet: {firstRow.IndirizzoGasnet}");
            
            if (string.IsNullOrEmpty(firstRow.OraInizioServizio))
            {
                Console.WriteLine("✓ Ora Inizio Servizio is empty (CORRECT)\n");
            }
            else
            {
                Console.WriteLine("✗ Ora Inizio Servizio is NOT empty (WRONG)\n");
            }
        }
        
        // Test 4: Open Excel and check fissi sheet
        Console.WriteLine("Test 4: Opening Excel file...");
        var workbook = excelManager.OpenWorkbook("2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx");
        var fissiSheet = excelManager.GetFissiSheet(workbook);
        Console.WriteLine("✓ Found fissi sheet");
        
        // Check row 2 of fissi sheet
        var row2Cell = fissiSheet.Worksheet.Cells[2, 1];
        Console.WriteLine($"  Fissi Row 2, Cell A2: '{row2Cell.Text}'");
        if (row2Cell.Text?.Trim().Equals("Data", StringComparison.OrdinalIgnoreCase) == true)
        {
            Console.WriteLine("✓ Row 2 contains 'Data' header (will be skipped)\n");
        }
        
        Console.WriteLine("\n=== TEST COMPLETE ===");
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
