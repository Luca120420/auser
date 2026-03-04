using System.Collections.Generic;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Interface for the data transformer service that converts CSV data into Excel format.
    /// Applies all transformation rules to service appointments.
    /// Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9
    /// </summary>
    public interface IDataTransformer
    {
        /// <summary>
        /// Transforms a list of service appointments into the Excel output format.
        /// Applies all transformation rules and returns the result with highlight information.
        /// </summary>
        /// <param name="appointments">The list of service appointments from the CSV file</param>
        /// <returns>A TransformationResult containing transformed rows and yellow highlight information</returns>
        TransformationResult Transform(List<ServiceAppointment> appointments);

        /// <summary>
        /// Transforms a list of service appointments into the enhanced Excel output format.
        /// Performs lookups against reference sheets and populates all 15 columns.
        /// Validates: Requirements 5.1, 5.2, 6.1, 6.2, 7.2, 8.2, 8.3
        /// </summary>
        /// <param name="appointments">The list of service appointments from the CSV file</param>
        /// <param name="lookupService">The lookup service for reference sheet operations</param>
        /// <returns>An EnhancedTransformationResult containing enhanced rows with lookup data</returns>
        EnhancedTransformationResult TransformEnhanced(List<ServiceAppointment> appointments, ILookupService lookupService);
    }
}
