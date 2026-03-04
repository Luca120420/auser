using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace AuserExcelTransformer.Models
{
    /// <summary>
    /// Represents the structure of volontari-auser.json file.
    /// </summary>
    public class VolunteerFileData
    {
        /// <summary>
        /// Dictionary mapping volunteer surnames to email addresses.
        /// JSON property name: "associates"
        /// </summary>
        [JsonPropertyName("associates")]
        public Dictionary<string, string> Associates { get; set; } = new Dictionary<string, string>();
    }
}
