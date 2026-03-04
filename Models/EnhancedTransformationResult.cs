using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace AuserExcelTransformer.Models
{
    /// <summary>
    /// Represents the result of transforming CSV data into enhanced Excel format.
    /// Contains the enhanced transformed rows with lookup data and metadata about yellow highlighting.
    /// </summary>
    public class EnhancedTransformationResult
    {
        /// <summary>
        /// Rows - The list of enhanced transformed data rows ready for Excel output
        /// Each row has been processed with lookups and all 15 columns populated
        /// </summary>
        [Required(ErrorMessage = "Rows è obbligatorio")]
        public List<EnhancedTransformedRow> Rows { get; set; } = new List<EnhancedTransformedRow>();

        /// <summary>
        /// YellowHighlightRows - List of row indices that should be highlighted in yellow
        /// These are rows where ATTIVITÀ contained "Accompag. con macchina attrezzata"
        /// Row indices are 1-based to match Excel row numbering (excluding header)
        /// </summary>
        [Required(ErrorMessage = "YellowHighlightRows è obbligatorio")]
        public List<int> YellowHighlightRows { get; set; } = new List<int>();
    }
}
