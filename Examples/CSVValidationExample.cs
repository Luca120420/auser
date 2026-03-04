using System;
using System.Collections.Generic;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Examples
{
    /// <summary>
    /// Example demonstrating how to use the CSV validation with detailed error information.
    /// This shows how the application controller can provide user-friendly error messages
    /// listing the specific missing columns.
    /// Validates: Requirements 2.3, 9.3
    /// </summary>
    public class CSVValidationExample
    {
        public static void DemonstrateValidation()
        {
            var parser = new CSVParser();
            string csvFilePath = "path/to/file.csv";

            // Use the detailed validation method
            List<string> missingColumns;
            bool isValid = parser.ValidateCSVStructure(csvFilePath, out missingColumns);

            if (!isValid)
            {
                if (missingColumns.Count > 0)
                {
                    // Create Italian error message listing missing columns
                    string columnList = string.Join(", ", missingColumns);
                    string errorMessage = $"Il file CSV non contiene le colonne richieste: {columnList}";
                    
                    Console.WriteLine(errorMessage);
                    // In the actual application, this would be displayed in the GUI
                }
                else
                {
                    Console.WriteLine("Impossibile leggere il file CSV.");
                }
            }
            else
            {
                Console.WriteLine("Il file CSV è valido e contiene tutte le colonne richieste.");
            }
        }
    }
}
