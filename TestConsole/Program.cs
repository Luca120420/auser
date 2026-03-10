using System;
using System.Linq;
using AuserExcelTransformer.Services;

class Program
{
    static void Main()
    {
        try
        {
            Console.WriteLine("Starting test run...");
            
            var csvParser = new CSVParser();
            var dateCalculator = new DateCalculator();
            var headerCalculator = new HeaderCalculator(dateCalculator);
            var transformationRulesEngine = new TransformationRulesEngine();
            var dataTransformer = new DataTransformer(transformationRulesEngine);
            var excelManager = new ExcelManager();
            
            // Parse CSV
            Console.WriteLine("Parsing CSV...");
            var appointments = csvParser.ParseCSV("168514-Estrazione_1770193162042.csv");
            Console.WriteLine($"Parsed {appointments.Count} appointments");
            
            // Transform data
            Console.WriteLine("Transforming data...");
            var result = dataTransformer.Transform(appointments);
            Console.WriteLine($"Transformed to {result.Rows.Count} rows");
            Console.WriteLine($"Yellow highlight rows: {result.YellowHighlightRows.Count}");
            
            // Print first row to verify mapping
            if (result.Rows.Count > 0)
            {
                var firstRow = result.Rows[0];
                Console.WriteLine("\nFirst row:");
                Console.WriteLine($"  Data: {firstRow.DataServizio}");
                Console.WriteLine($"  Ora Inizio Servizio: {firstRow.OraInizioServizio}");
                Console.WriteLine($"  Assistito: {firstRow.Assistito}");
                Console.WriteLine($"  Destinazione: {firstRow.Destinazione}");
                Console.WriteLine($"  Note: {firstRow.Note}");
                Console.WriteLine($"  Auto: {firstRow.Auto}");
                Console.WriteLine($"  Volontario: {firstRow.Volontario}");
                Console.WriteLine($"  Arrivo: {firstRow.Arrivo}");
                Console.WriteLine($"  Indirizzo Partenza: {firstRow.IndirizzoPartenza}");
                Console.WriteLine($"  Comune Partenza: {firstRow.ComunePartenza}");
                Console.WriteLine($"  Indirizzo Gasnet: {firstRow.IndirizzoGasnet}");
            }
            
            // Open Excel
            Console.WriteLine("\nOpening Excel file...");
            var workbook = excelManager.OpenWorkbook("2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx");
            
            // Get sheet info
            var sheetNames = excelManager.GetSheetNames(workbook);
            Console.WriteLine($"Found {sheetNames.Count} sheets");
            
            int nextSheetNumber = excelManager.GetNextSheetNumber(sheetNames);
            Console.WriteLine($"Next sheet number: {nextSheetNumber}");
            
            // Get fissi sheet
            var fissiSheet = excelManager.GetFissiSheet(workbook);
            Console.WriteLine("Found fissi sheet");
            
            // Read previous header
            int lastSheetNumber = nextSheetNumber - 1;
            var lastSheet = excelManager.GetSheetByName(workbook, lastSheetNumber.ToString());
            string previousHeader = excelManager.ReadHeader(lastSheet);
            Console.WriteLine($"Previous header: {previousHeader}");
            
            // Parse header
            var previousHeaderInfo = headerCalculator.ParseHeader(previousHeader);
            Console.WriteLine($"Previous Monday: {previousHeaderInfo.MondayDate:yyyy-MM-dd}");
            
            // Calculate next Monday
            var nextMondayDate = previousHeaderInfo.MondayDate.AddDays(7);
            Console.WriteLine($"Next Monday: {nextMondayDate:yyyy-MM-dd}");
            
            // Create new sheet
            Console.WriteLine("\nCreating new sheet...");
            var newSheet = excelManager.CreateNewSheet(workbook, nextSheetNumber);
            
            // Write header and column headers
            Console.WriteLine("Writing headers...");
            excelManager.WriteHeaderRow(newSheet, nextMondayDate);
            excelManager.WriteColumnHeaders(newSheet);
            
            // Write data
            Console.WriteLine("Writing data rows...");
            excelManager.WriteDataRows(newSheet, result.Rows, 3);
            
            // Append fissi data
            int fissiStartRow = 3 + result.Rows.Count;
            Console.WriteLine($"Appending fissi data at row {fissiStartRow}...");
            excelManager.AppendFissiData(newSheet, fissiSheet, fissiStartRow, DateTime.Now);
            
            // Apply yellow highlighting
            var adjustedHighlightRows = result.YellowHighlightRows.Select(r => r + 2).ToList();
            Console.WriteLine($"Applying yellow highlight to {adjustedHighlightRows.Count} rows...");
            excelManager.ApplyYellowHighlight(newSheet, adjustedHighlightRows);
            
            // Enable AutoFilter
            Console.WriteLine("Enabling AutoFilter...");
            excelManager.EnableAutoFilter(newSheet);
            
            // Save
            string outputPath = "test_output.xlsx";
            Console.WriteLine($"\nSaving to {outputPath}...");
            excelManager.SaveWorkbook(workbook, outputPath);
            
            Console.WriteLine("\nSUCCESS! File saved.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\nERROR: {ex.Message}");
            Console.WriteLine($"Stack: {ex.StackTrace}");
        }
    }
}
