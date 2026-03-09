using System;
using System.Linq;
using AuserExcelTransformer.Services;

class TestYellowHighlight
{
    static void Main()
    {
        Console.WriteLine("=== Test Evidenziazione Gialla ===\n");
        
        // Parse CSV
        var csvParser = new CSVParser();
        var appointments = csvParser.ParseCSV("168514-Estrazione_1772702766454.csv");
        
        Console.WriteLine($"Totale righe CSV: {appointments.Count}");
        
        // Conta righe con "Accompag. con macchina attrezzata"
        var yellowRows = appointments
            .Where(a => !string.IsNullOrEmpty(a.Attivita) && 
                       a.Attivita.Contains("Accompag. con macchina attrezzata"))
            .ToList();
        
        Console.WriteLine($"Righe con 'Accompag. con macchina attrezzata': {yellowRows.Count}\n");
        
        // Mostra le righe che dovrebbero essere evidenziate
        Console.WriteLine("Dettagli righe da evidenziare:");
        foreach (var row in yellowRows)
        {
            Console.WriteLine($"  - {row.DataServizio} | {row.CognomeAssistito} {row.NomeAssistito} | Stato: {row.DescrizioneStatoServizio}");
        }
        
        // Test con DataTransformer
        Console.WriteLine("\n=== Test con DataTransformer ===\n");
        var rulesEngine = new TransformationRulesEngine();
        var dataTransformer = new DataTransformer(rulesEngine);
        var result = dataTransformer.Transform(appointments);
        
        Console.WriteLine($"Righe trasformate: {result.Rows.Count}");
        Console.WriteLine($"Righe da evidenziare in giallo: {result.YellowHighlightRows.Count}");
        
        if (result.YellowHighlightRows.Count > 0)
        {
            Console.WriteLine("\nIndici righe da evidenziare:");
            foreach (var rowIndex in result.YellowHighlightRows)
            {
                Console.WriteLine($"  - Riga {rowIndex}");
            }
        }
        
        // Verifica che le righe ANNULLATO non siano incluse
        var cancelledYellow = yellowRows.Where(a => 
            !string.IsNullOrEmpty(a.DescrizioneStatoServizio) && 
            a.DescrizioneStatoServizio.Equals("ANNULLATO", StringComparison.OrdinalIgnoreCase))
            .Count();
        
        Console.WriteLine($"\nRighe 'Accompag. con macchina attrezzata' ANNULLATE (non evidenziate): {cancelledYellow}");
        Console.WriteLine($"Righe 'Accompag. con macchina attrezzata' da evidenziare (escluse ANNULLATE): {yellowRows.Count - cancelledYellow}");
        
        Console.WriteLine("\n✓ Test completato!");
    }
}
