using System;
using System.Collections.Generic;
using System.IO;
using NUnit.Framework;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for VolunteerManager class.
    /// Tests specific examples, edge cases, and error conditions.
    /// Validates: Requirements 1.5, 8.12
    /// </summary>
    [TestFixture]
    public class VolunteerManagerTests
    {
        private VolunteerManager _volunteerManager = null!;

        [SetUp]
        public void Setup()
        {
            _volunteerManager = new VolunteerManager();
        }

        #region Invalid JSON Rejection Tests (Property 3)

        /// <summary>
        /// Property 3: Invalid JSON Rejection
        /// For any malformed JSON file, attempting to load it as a volunteer file
        /// should result in an error and the file should be rejected.
        /// **Validates: Requirements 1.5**
        /// </summary>
        [Test]
        public void LoadVolunteers_WithMalformedJson_ShouldThrowInvalidOperationException()
        {
            // Arrange
            var tempFile = Path.GetTempFileName();
            try
            {
                File.WriteAllText(tempFile, "{ invalid json syntax }");

                // Act & Assert
                var ex = Assert.Throws<InvalidOperationException>(() =>
                    _volunteerManager.LoadVolunteers(tempFile));

                Assert.That(ex.Message, Does.Contain("Failed to parse JSON file"));
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }

        [Test]
        public void LoadVolunteers_WithMissingAssociatesProperty_ShouldReturnEmptyDictionary()
        {
            // Arrange
            var tempFile = Path.GetTempFileName();
            try
            {
                // Valid JSON but missing "associates" property
                // System.Text.Json deserializes this as an object with empty Associates dictionary
                File.WriteAllText(tempFile, "{ \"otherProperty\": \"value\" }");

                // Act
                var volunteers = _volunteerManager.LoadVolunteers(tempFile);

                // Assert - Should return empty dictionary when associates property is missing
                Assert.That(volunteers, Is.Not.Null);
                Assert.That(volunteers.Count, Is.EqualTo(0));
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }

        [Test]
        public void LoadVolunteers_WithNullAssociatesProperty_ShouldThrowInvalidOperationException()
        {
            // Arrange
            var tempFile = Path.GetTempFileName();
            try
            {
                // Valid JSON but associates is null
                File.WriteAllText(tempFile, "{ \"associates\": null }");

                // Act & Assert
                var ex = Assert.Throws<InvalidOperationException>(() =>
                    _volunteerManager.LoadVolunteers(tempFile));

                Assert.That(ex.Message, Does.Contain("Invalid JSON structure"));
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }

        [Test]
        public void LoadVolunteers_WithEmptyJson_ShouldThrowInvalidOperationException()
        {
            // Arrange
            var tempFile = Path.GetTempFileName();
            try
            {
                File.WriteAllText(tempFile, "");

                // Act & Assert
                var ex = Assert.Throws<InvalidOperationException>(() =>
                    _volunteerManager.LoadVolunteers(tempFile));

                Assert.That(ex.Message, Does.Contain("Failed to parse JSON file"));
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }

        [Test]
        public void LoadVolunteers_WithIncompleteBraces_ShouldThrowInvalidOperationException()
        {
            // Arrange
            var tempFile = Path.GetTempFileName();
            try
            {
                File.WriteAllText(tempFile, "{ \"associates\": { \"Rossi\": \"rossi@example.com\" ");

                // Act & Assert
                var ex = Assert.Throws<InvalidOperationException>(() =>
                    _volunteerManager.LoadVolunteers(tempFile));

                Assert.That(ex.Message, Does.Contain("Failed to parse JSON file"));
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }

        #endregion

        #region Empty Surname Rejection Tests

        /// <summary>
        /// Tests that AddVolunteer rejects empty surnames.
        /// **Validates: Requirements 8.12**
        /// </summary>
        [Test]
        public void AddVolunteer_WithEmptySurname_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var email = "test@example.com";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer("", email, volunteers));

            Assert.That(ex.Message, Does.Contain("Surname cannot be empty"));
            Assert.That(ex.ParamName, Is.EqualTo("surname"));
        }

        [Test]
        public void AddVolunteer_WithWhitespaceSurname_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var email = "test@example.com";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer("   ", email, volunteers));

            Assert.That(ex.Message, Does.Contain("Surname cannot be empty"));
            Assert.That(ex.ParamName, Is.EqualTo("surname"));
        }

        [Test]
        public void AddVolunteer_WithNullSurname_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var email = "test@example.com";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer(null!, email, volunteers));

            Assert.That(ex.Message, Does.Contain("Surname cannot be empty"));
            Assert.That(ex.ParamName, Is.EqualTo("surname"));
        }

        [Test]
        public void AddVolunteer_WithTabsAndSpacesSurname_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var email = "test@example.com";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer("\t  \t", email, volunteers));

            Assert.That(ex.Message, Does.Contain("Surname cannot be empty"));
            Assert.That(ex.ParamName, Is.EqualTo("surname"));
        }

        #endregion

        #region Invalid Email Rejection Tests (Property 22)

        /// <summary>
        /// Property 22: Invalid Contact Rejection
        /// For any contact with an empty surname or invalid email format,
        /// attempting to add the contact should result in an error message
        /// and the contact should not be added to the list.
        /// **Validates: Requirements 8.12**
        /// </summary>
        [Test]
        public void AddVolunteer_WithInvalidEmail_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var surname = "Rossi";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer(surname, "invalid-email", volunteers));

            Assert.That(ex.Message, Does.Contain("Email address is not valid"));
            Assert.That(ex.ParamName, Is.EqualTo("email"));
        }

        [Test]
        public void AddVolunteer_WithEmptyEmail_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var surname = "Rossi";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer(surname, "", volunteers));

            Assert.That(ex.Message, Does.Contain("Email address is not valid"));
            Assert.That(ex.ParamName, Is.EqualTo("email"));
        }

        [Test]
        public void AddVolunteer_WithNullEmail_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var surname = "Rossi";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer(surname, null!, volunteers));

            Assert.That(ex.Message, Does.Contain("Email address is not valid"));
            Assert.That(ex.ParamName, Is.EqualTo("email"));
        }

        [Test]
        public void AddVolunteer_WithWhitespaceEmail_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var surname = "Rossi";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer(surname, "   ", volunteers));

            Assert.That(ex.Message, Does.Contain("Email address is not valid"));
            Assert.That(ex.ParamName, Is.EqualTo("email"));
        }

        [Test]
        public void AddVolunteer_WithEmailMissingAtSymbol_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var surname = "Rossi";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer(surname, "testexample.com", volunteers));

            Assert.That(ex.Message, Does.Contain("Email address is not valid"));
        }

        [Test]
        public void AddVolunteer_WithEmailMissingDomain_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var surname = "Rossi";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer(surname, "test@", volunteers));

            Assert.That(ex.Message, Does.Contain("Email address is not valid"));
        }

        [Test]
        public void AddVolunteer_WithEmailMissingLocalPart_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var surname = "Rossi";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer(surname, "@example.com", volunteers));

            Assert.That(ex.Message, Does.Contain("Email address is not valid"));
        }

        [Test]
        public void AddVolunteer_WithMultipleAtSymbols_ShouldThrowArgumentException()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var surname = "Rossi";

            // Act & Assert
            var ex = Assert.Throws<ArgumentException>(() =>
                _volunteerManager.AddVolunteer(surname, "test@@example.com", volunteers));

            Assert.That(ex.Message, Does.Contain("Email address is not valid"));
        }

        #endregion

        #region IsValidEmail Tests

        [Test]
        public void IsValidEmail_WithValidEmail_ShouldReturnTrue()
        {
            // Act & Assert
            Assert.That(_volunteerManager.IsValidEmail("test@example.com"), Is.True);
            Assert.That(_volunteerManager.IsValidEmail("user.name@domain.co.uk"), Is.True);
            Assert.That(_volunteerManager.IsValidEmail("first.last@subdomain.example.org"), Is.True);
        }

        [Test]
        public void IsValidEmail_WithInvalidEmail_ShouldReturnFalse()
        {
            // Act & Assert
            Assert.That(_volunteerManager.IsValidEmail("invalid"), Is.False);
            Assert.That(_volunteerManager.IsValidEmail("@example.com"), Is.False);
            Assert.That(_volunteerManager.IsValidEmail("test@"), Is.False);
            Assert.That(_volunteerManager.IsValidEmail("test@@example.com"), Is.False);
            Assert.That(_volunteerManager.IsValidEmail(""), Is.False);
            Assert.That(_volunteerManager.IsValidEmail(null!), Is.False);
            Assert.That(_volunteerManager.IsValidEmail("   "), Is.False);
        }

        #endregion

        #region Successful Operations Tests

        [Test]
        public void AddVolunteer_WithValidData_ShouldAddToList()
        {
            // Arrange
            var volunteers = new Dictionary<string, string>();
            var surname = "Rossi";
            var email = "rossi@example.com";

            // Act
            _volunteerManager.AddVolunteer(surname, email, volunteers);

            // Assert
            Assert.That(volunteers.Count, Is.EqualTo(1));
            Assert.That(volunteers.ContainsKey(surname), Is.True);
            Assert.That(volunteers[surname], Is.EqualTo(email));
        }

        [Test]
        public void LoadVolunteers_WithValidFile_ShouldReturnVolunteers()
        {
            // Arrange
            var tempFile = Path.GetTempFileName();
            try
            {
                var validJson = @"{
                    ""associates"": {
                        ""Rossi"": ""rossi@example.com"",
                        ""Bianchi"": ""bianchi@example.com""
                    }
                }";
                File.WriteAllText(tempFile, validJson);

                // Act
                var volunteers = _volunteerManager.LoadVolunteers(tempFile);

                // Assert
                Assert.That(volunteers.Count, Is.EqualTo(2));
                Assert.That(volunteers["Rossi"], Is.EqualTo("rossi@example.com"));
                Assert.That(volunteers["Bianchi"], Is.EqualTo("bianchi@example.com"));
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }

        [Test]
        public void LoadVolunteers_WithNonExistentFile_ShouldThrowFileNotFoundException()
        {
            // Arrange
            var nonExistentFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".json");

            // Act & Assert
            Assert.Throws<FileNotFoundException>(() =>
                _volunteerManager.LoadVolunteers(nonExistentFile));
        }

        #endregion
    }
}
