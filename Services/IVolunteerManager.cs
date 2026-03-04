using System;
using System.Collections.Generic;
using System.IO;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Manages volunteer contact information with add, delete, and persistence operations.
    /// </summary>
    public interface IVolunteerManager
    {
        /// <summary>
        /// Loads volunteer contacts from the specified JSON file path.
        /// </summary>
        /// <param name="filePath">Path to volontari-auser.json file</param>
        /// <returns>Dictionary mapping surname to email address</returns>
        /// <exception cref="FileNotFoundException">If file does not exist</exception>
        /// <exception cref="InvalidOperationException">If JSON is malformed</exception>
        Dictionary<string, string> LoadVolunteers(string filePath);
        
        /// <summary>
        /// Saves volunteer contacts to the specified JSON file path.
        /// </summary>
        /// <param name="filePath">Path to volontari-auser.json file</param>
        /// <param name="volunteers">Dictionary mapping surname to email address</param>
        /// <exception cref="IOException">If file cannot be written</exception>
        void SaveVolunteers(string filePath, Dictionary<string, string> volunteers);
        
        /// <summary>
        /// Adds a new volunteer contact.
        /// </summary>
        /// <param name="surname">Volunteer surname</param>
        /// <param name="email">Volunteer email address</param>
        /// <param name="volunteers">Current volunteer dictionary</param>
        /// <exception cref="ArgumentException">If surname is empty or email is invalid</exception>
        void AddVolunteer(string surname, string email, Dictionary<string, string> volunteers);
        
        /// <summary>
        /// Removes a volunteer contact by surname.
        /// </summary>
        /// <param name="surname">Volunteer surname to remove</param>
        /// <param name="volunteers">Current volunteer dictionary</param>
        void RemoveVolunteer(string surname, Dictionary<string, string> volunteers);
        
        /// <summary>
        /// Validates email address format.
        /// </summary>
        /// <param name="email">Email address to validate</param>
        /// <returns>True if valid, false otherwise</returns>
        bool IsValidEmail(string email);
    }
}
