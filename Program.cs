using System;
using System.Linq;
using System.Windows.Forms;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.UI;
using AuserExcelTransformer.Examples;

namespace AuserExcelTransformer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// Validates: Requirements 10.1, 10.2
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            // Debug: Log arguments to a file to see what's being passed
            if (args.Length > 0)
            {
                try
                {
                    System.IO.File.WriteAllText("debug_args.txt", 
                        $"Args count: {args.Length}\n" + 
                        string.Join("\n", args.Select((a, i) => $"Arg[{i}]: {a}")));
                }
                catch { }
            }
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // Set Italian culture for the application
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("it-IT");
            
            // Initialize services
            var csvParser = new CSVParser();
            var dateCalculator = new DateCalculator();
            var headerCalculator = new HeaderCalculator(dateCalculator);
            var transformationRulesEngine = new TransformationRulesEngine();
            var dataTransformer = new DataTransformer(transformationRulesEngine);
            var excelManager = new ExcelManager();
            
            // Create a simple wrapper to handle the circular dependency
            GUIWrapper guiWrapper = new GUIWrapper();
            
            // Create controller
            var controller = new ApplicationController(
                guiWrapper,
                csvParser,
                excelManager,
                dataTransformer,
                headerCalculator
            );
            
            // Create the actual form and set it in the wrapper
            var mainForm = new MainForm(controller);
            guiWrapper.SetGUI(mainForm);
            
            // Run the application
            Application.Run(mainForm);
        }
    }
    
    /// <summary>
    /// Wrapper class to handle circular dependency between MainForm and ApplicationController.
    /// </summary>
    internal class GUIWrapper : IGUI
    {
        private IGUI? _actualGUI;
        
        public void SetGUI(IGUI gui)
        {
            _actualGUI = gui;
        }
        
        public void ShowWindow() => _actualGUI?.ShowWindow();
        public string? SelectCSVFile() => _actualGUI?.SelectCSVFile();
        public string? SelectExcelFile() => _actualGUI?.SelectExcelFile();
        public void DisplaySelectedCSVPath(string path) => _actualGUI?.DisplaySelectedCSVPath(path);
        public void DisplaySelectedExcelPath(string path) => _actualGUI?.DisplaySelectedExcelPath(path);
        public void EnableProcessButton(bool enabled) => _actualGUI?.EnableProcessButton(enabled);
        public void EnableDownloadButton(bool enabled) => _actualGUI?.EnableDownloadButton(enabled);
        public void ShowErrorMessage(string message) => _actualGUI?.ShowErrorMessage(message);
        public void ShowSuccessMessage(string message) => _actualGUI?.ShowSuccessMessage(message);
        public string? GetSaveFilePath() => _actualGUI?.GetSaveFilePath();
    }
}
