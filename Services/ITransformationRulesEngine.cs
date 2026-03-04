using System.Collections.Generic;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Interface for the transformation rules engine that applies all data transformation rules.
    /// Implements the 11 specific transformation rules for converting CSV data to Excel format.
    /// Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9
    /// </summary>
    public interface ITransformationRulesEngine
    {
        /// <summary>
        /// Transforms a list of service appointments according to all transformation rules.
        /// </summary>
        /// <param name="appointments">The list of service appointments from the CSV file</param>
        /// <returns>A TransformationResult containing transformed rows and yellow highlight information</returns>
        TransformationResult Transform(List<ServiceAppointment> appointments);
    }
}
