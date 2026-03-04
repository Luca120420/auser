using OfficeOpenXml;

namespace AuserExcelTransformer.Models
{
    /// <summary>
    /// Wrapper class around EPPlus ExcelPackage to provide abstraction for Excel workbook operations.
    /// This allows for easier testing and potential library swapping in the future.
    /// </summary>
    public class ExcelWorkbook
    {
        /// <summary>
        /// The underlying EPPlus ExcelPackage object.
        /// </summary>
        public ExcelPackage Package { get; }

        /// <summary>
        /// Initializes a new instance of the ExcelWorkbook class.
        /// </summary>
        /// <param name="package">The EPPlus ExcelPackage to wrap</param>
        public ExcelWorkbook(ExcelPackage package)
        {
            Package = package;
        }

        /// <summary>
        /// Gets the workbook from the package.
        /// </summary>
        public ExcelWorkbook Workbook => this;
    }
}
