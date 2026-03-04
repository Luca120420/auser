using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Manages volunteer contact information with add, delete, and persistence operations.
    /// Validates: Requirements 1.4, 1.5, 1.6, 8.4, 8.10, 8.12
    /// </summary>
    public class VolunteerManager : IVolunteerManager
    {
        /// <summary>
        /// Loads volunteer contacts from the specified JSON file path.
        /// </summary>
        /// <param name="filePath">Path to volontari-auser.json file</param>
        /// <returns>Dictionary mapping surname to email address</returns>
        /// <exception cref="FileNotFoundException">If file does not exist</exception>
        /// <exception cref="InvalidOperationException">If JSON is malformed</exception>
        public Dictionary<string, string> LoadVolunteers(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Volunteer file not found: {filePath}", filePath);
            }

            try
            {
                string jsonContent = File.ReadAllText(filePath);
                var volunteerFileData = JsonSerializer.Deserialize<VolunteerFileData>(jsonContent);

                if (volunteerFileData == null || volunteerFileData.Associates == null)
                {
                    throw new InvalidOperationException("Invalid JSON structure: missing 'associates' property");
                }

                return volunteerFileData.Associates;
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException($"Failed to parse JSON file: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Saves volunteer contacts to the specified JSON file path.
        /// </summary>
        /// <param name="filePath">Path to volontari-auser.json file</param>
        /// <param name="volunteers">Dictionary mapping surname to email address</param>
        /// <exception cref="IOException">If file cannot be written</exception>
        public void SaveVolunteers(string filePath, Dictionary<string, string> volunteers)
        {
            try
            {
                var volunteerFileData = new VolunteerFileData
                {
                    Associates = volunteers ?? new Dictionary<string, string>()
                };

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                string jsonContent = JsonSerializer.Serialize(volunteerFileData, options);
                File.WriteAllText(filePath, jsonContent);
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException || ex is DirectoryNotFoundException)
            {
                throw new IOException($"Failed to write volunteer file: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Adds a new volunteer contact.
        /// </summary>
        /// <param name="surname">Volunteer surname</param>
        /// <param name="email">Volunteer email address</param>
        /// <param name="volunteers">Current volunteer dictionary</param>
        /// <exception cref="ArgumentException">If surname is empty or email is invalid</exception>
        public void AddVolunteer(string surname, string email, Dictionary<string, string> volunteers)
        {
            if (string.IsNullOrWhiteSpace(surname))
            {
                throw new ArgumentException("Surname cannot be empty or whitespace", nameof(surname));
            }

            if (!IsValidEmail(email))
            {
                throw new ArgumentException("Email address is not valid", nameof(email));
            }

            volunteers[surname] = email;
        }

        /// <summary>
        /// Removes a volunteer contact by surname.
        /// </summary>
        /// <param name="surname">Volunteer surname to remove</param>
        /// <param name="volunteers">Current volunteer dictionary</param>
        public void RemoveVolunteer(string surname, Dictionary<string, string> volunteers)
        {
            volunteers.Remove(surname);
        }

        /// <summary>
        /// Validates email address format.
        /// </summary>
        /// <param name="email">Email address to validate</param>
        /// <returns>True if valid, false otherwise</returns>
        public bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                return false;
            }

            try
            {
                var mailAddress = new System.Net.Mail.MailAddress(email);
                return mailAddress.Address == email;
            }
            catch
            {
                return false;
            }
        }
    }
}
