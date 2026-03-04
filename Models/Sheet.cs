using OfficeOpenXml;

namespace AuserExcelTransformer.Models
{
    /// <summary>
    /// Wrapper class around EPPlus ExcelWorksheet to provide abstraction for Excel sheet operations.
    /// This allows for easier testing and potential library swapping in the future.
    /// </summary>
    public class Sheet
    {
        /// <summary>
        /// The underlying EPPlus ExcelWorksheet object.
        /// </summary>
        public ExcelWorksheet Worksheet { get; }

        /// <summary>
        /// Initializes a new instance of the Sheet class.
        /// </summary>
        /// <param name="worksheet">The EPPlus ExcelWorksheet to wrap</param>
        public Sheet(ExcelWorksheet worksheet)
        {
            Worksheet = worksheet;
        }

        /// <summary>
        /// Gets the name of the sheet.
        /// </summary>
        public string Name => Worksheet.Name;
    }
}
