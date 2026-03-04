using System.Collections.Generic;

namespace AuserExcelTransformer.Models;

/// <summary>
/// Represents a volunteer's assigned rows from Excel data.
/// </summary>
public class VolunteerAssignment
{
    /// <summary>
    /// Volunteer surname.
    /// </summary>
    public string Surname { get; set; } = string.Empty;
    
    /// <summary>
    /// Volunteer email address.
    /// </summary>
    public string Email { get; set; } = string.Empty;
    
    /// <summary>
    /// List of assigned rows (each row is a dictionary of column name to value).
    /// </summary>
    public List<Dictionary<string, string>> AssignedRows { get; set; } = new List<Dictionary<string, string>>();
}
