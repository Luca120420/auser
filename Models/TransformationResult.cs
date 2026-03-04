using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace AuserExcelTransformer.Models
{
    /// <summary>
    /// Represents the result of transforming CSV data into Excel format.
    /// Contains the transformed rows and metadata about which rows need yellow highlighting.
    /// Validates: Requirements 4.1
    /// </summary>
    public class TransformationResult
    {
        /// <summary>
        /// Rows - The list of transformed data rows ready for Excel output
        /// Each row has been processed according to all transformation rules
        /// </summary>
        [Required(ErrorMessage = "Rows è obbligatorio")]
        public List<TransformedRow> Rows { get; set; } = new List<TransformedRow>();

        /// <summary>
        /// YellowHighlightRows - List of row indices that should be highlighted in yellow
        /// These are rows where ATTIVITÀ contained "Accompag. con macchina attrezzata"
        /// Row indices are 1-based to match Excel row numbering (excluding header)
        /// </summary>
        [Required(ErrorMessage = "YellowHighlightRows è obbligatorio")]
        public List<int> YellowHighlightRows { get; set; } = new List<int>();
    }
}
