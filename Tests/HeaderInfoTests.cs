using System;
using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for the HeaderInfo model class.
    /// Validates: Requirements 5.1, 5.2
    /// </summary>
    [TestFixture]
    public class HeaderInfoTests
    {
        [Test]
        public void HeaderInfo_CanBeInstantiated()
        {
            // Arrange & Act
            var headerInfo = new HeaderInfo();

            // Assert
            Assert.That(headerInfo, Is.Not.Null);
        }

        [Test]
        public void HeaderInfo_AllPropertiesCanBeSet()
        {
            // Arrange
            var mondayDate = new DateTime(2024, 1, 29);
            var sundayDate = new DateTime(2024, 2, 4);
            var weekNumber = 5;
            var referente = "Mario Rossi - 333-1234567";

            // Act
            var headerInfo = new HeaderInfo
            {
                MondayDate = mondayDate,
                SundayDate = sundayDate,
                WeekNumber = weekNumber,
                Referente = referente
            };

            // Assert
            Assert.That(headerInfo.MondayDate, Is.EqualTo(mondayDate));
            Assert.That(headerInfo.SundayDate, Is.EqualTo(sundayDate));
            Assert.That(headerInfo.WeekNumber, Is.EqualTo(weekNumber));
            Assert.That(headerInfo.Referente, Is.EqualTo(referente));
        }

        [Test]
        public void HeaderInfo_MondayDate_CanBeSet()
        {
            // Arrange
            var mondayDate = new DateTime(2024, 1, 29);
            var headerInfo = new HeaderInfo
            {
                MondayDate = mondayDate,
                SundayDate = new DateTime(2024, 2, 4),
                WeekNumber = 5,
                Referente = "Test Referente"
            };

            // Act & Assert
            Assert.That(headerInfo.MondayDate, Is.EqualTo(mondayDate));
            Assert.That(headerInfo.MondayDate.Year, Is.EqualTo(2024));
            Assert.That(headerInfo.MondayDate.Month, Is.EqualTo(1));
            Assert.That(headerInfo.MondayDate.Day, Is.EqualTo(29));
        }

        [Test]
        public void HeaderInfo_SundayDate_CanBeSet()
        {
            // Arrange
            var sundayDate = new DateTime(2024, 2, 4);
            var headerInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2024, 1, 29),
                SundayDate = sundayDate,
                WeekNumber = 5,
                Referente = "Test Referente"
            };

            // Act & Assert
            Assert.That(headerInfo.SundayDate, Is.EqualTo(sundayDate));
            Assert.That(headerInfo.SundayDate.Year, Is.EqualTo(2024));
            Assert.That(headerInfo.SundayDate.Month, Is.EqualTo(2));
            Assert.That(headerInfo.SundayDate.Day, Is.EqualTo(4));
        }

        [Test]
        public void HeaderInfo_WeekNumber_IsRequired()
        {
            // Arrange
            var headerInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2024, 1, 29),
                SundayDate = new DateTime(2024, 2, 4),
                // WeekNumber is default (0)
                Referente = "Test Referente"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(headerInfo);
            var isValid = Validator.TryValidateObject(headerInfo, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("WeekNumber")), Is.True);
        }

        [Test]
        public void HeaderInfo_WeekNumber_MustBeInValidRange()
        {
            // Arrange
            var headerInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2024, 1, 29),
                SundayDate = new DateTime(2024, 2, 4),
                WeekNumber = 54, // Invalid: exceeds maximum of 53
                Referente = "Test Referente"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(headerInfo);
            var isValid = Validator.TryValidateObject(headerInfo, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("WeekNumber")), Is.True);
        }

        [Test]
        public void HeaderInfo_WeekNumber_CannotBeZero()
        {
            // Arrange
            var headerInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2024, 1, 29),
                SundayDate = new DateTime(2024, 2, 4),
                WeekNumber = 0, // Invalid: below minimum of 1
                Referente = "Test Referente"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(headerInfo);
            var isValid = Validator.TryValidateObject(headerInfo, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("WeekNumber")), Is.True);
        }

        [Test]
        public void HeaderInfo_WeekNumber_AcceptsMinimumValue()
        {
            // Arrange
            var headerInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2024, 1, 1),
                SundayDate = new DateTime(2024, 1, 7),
                WeekNumber = 1, // Minimum valid value
                Referente = "Test Referente"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(headerInfo);
            var isValid = Validator.TryValidateObject(headerInfo, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
            Assert.That(validationResults.Count, Is.EqualTo(0));
        }

        [Test]
        public void HeaderInfo_WeekNumber_AcceptsMaximumValue()
        {
            // Arrange
            var headerInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2024, 12, 23),
                SundayDate = new DateTime(2024, 12, 29),
                WeekNumber = 53, // Maximum valid value
                Referente = "Test Referente"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(headerInfo);
            var isValid = Validator.TryValidateObject(headerInfo, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
            Assert.That(validationResults.Count, Is.EqualTo(0));
        }

        [Test]
        public void HeaderInfo_Referente_IsRequired()
        {
            // Arrange
            var headerInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2024, 1, 29),
                SundayDate = new DateTime(2024, 2, 4),
                WeekNumber = 5
                // Referente is empty (default value)
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(headerInfo);
            var isValid = Validator.TryValidateObject(headerInfo, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("Referente")), Is.True);
        }

        [Test]
        public void HeaderInfo_AllRequiredFieldsPresent_IsValid()
        {
            // Arrange
            var headerInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2024, 1, 29),
                SundayDate = new DateTime(2024, 2, 4),
                WeekNumber = 5,
                Referente = "Inserire nome e numero di telefono del referente"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(headerInfo);
            var isValid = Validator.TryValidateObject(headerInfo, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
            Assert.That(validationResults.Count, Is.EqualTo(0));
        }

        [Test]
        public void HeaderInfo_PreservesItalianCharacters()
        {
            // Arrange
            var headerInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2024, 1, 29),
                SundayDate = new DateTime(2024, 2, 4),
                WeekNumber = 5,
                Referente = "Nicolò Àgostini - 333-1234567"
            };

            // Act & Assert
            Assert.That(headerInfo.Referente, Is.EqualTo("Nicolò Àgostini - 333-1234567"));
        }

        [Test]
        public void HeaderInfo_DateRange_SpansOneWeek()
        {
            // Arrange
            var mondayDate = new DateTime(2024, 1, 29);
            var sundayDate = new DateTime(2024, 2, 4);

            var headerInfo = new HeaderInfo
            {
                MondayDate = mondayDate,
                SundayDate = sundayDate,
                WeekNumber = 5,
                Referente = "Test Referente"
            };

            // Act
            var daysDifference = (headerInfo.SundayDate - headerInfo.MondayDate).Days;

            // Assert
            Assert.That(daysDifference, Is.EqualTo(6)); // Monday to Sunday is 6 days
        }

        [Test]
        public void HeaderInfo_SupportsYearBoundaryTransition()
        {
            // Arrange - Week spanning from December to January
            var mondayDate = new DateTime(2023, 12, 25);
            var sundayDate = new DateTime(2023, 12, 31);

            var headerInfo = new HeaderInfo
            {
                MondayDate = mondayDate,
                SundayDate = sundayDate,
                WeekNumber = 52,
                Referente = "Test Referente"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(headerInfo);
            var isValid = Validator.TryValidateObject(headerInfo, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
            Assert.That(headerInfo.MondayDate.Year, Is.EqualTo(2023));
            Assert.That(headerInfo.SundayDate.Year, Is.EqualTo(2023));
        }

        [Test]
        public void HeaderInfo_SupportsMonthBoundaryTransition()
        {
            // Arrange - Week spanning from January to February
            var mondayDate = new DateTime(2024, 1, 29);
            var sundayDate = new DateTime(2024, 2, 4);

            var headerInfo = new HeaderInfo
            {
                MondayDate = mondayDate,
                SundayDate = sundayDate,
                WeekNumber = 5,
                Referente = "Test Referente"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(headerInfo);
            var isValid = Validator.TryValidateObject(headerInfo, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
            Assert.That(headerInfo.MondayDate.Month, Is.EqualTo(1));
            Assert.That(headerInfo.SundayDate.Month, Is.EqualTo(2));
        }

        [Test]
        public void HeaderInfo_DefaultReferenteText()
        {
            // Arrange
            var defaultReferente = "Inserire nome e numero di telefono del referente";
            var headerInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2024, 1, 29),
                SundayDate = new DateTime(2024, 2, 4),
                WeekNumber = 5,
                Referente = defaultReferente
            };

            // Act & Assert - Verify the default referente text matches requirement 5.7
            Assert.That(headerInfo.Referente, Is.EqualTo(defaultReferente));
        }

        [Test]
        public void HeaderInfo_HasCorrectPropertyTypes()
        {
            // Arrange
            var headerInfo = new HeaderInfo();
            var properties = typeof(HeaderInfo).GetProperties();

            // Act & Assert - Verify property types
            var mondayDateProp = properties.FirstOrDefault(p => p.Name == "MondayDate");
            var sundayDateProp = properties.FirstOrDefault(p => p.Name == "SundayDate");
            var weekNumberProp = properties.FirstOrDefault(p => p.Name == "WeekNumber");
            var referenteProp = properties.FirstOrDefault(p => p.Name == "Referente");

            Assert.That(mondayDateProp, Is.Not.Null);
            Assert.That(mondayDateProp!.PropertyType, Is.EqualTo(typeof(DateTime)));

            Assert.That(sundayDateProp, Is.Not.Null);
            Assert.That(sundayDateProp!.PropertyType, Is.EqualTo(typeof(DateTime)));

            Assert.That(weekNumberProp, Is.Not.Null);
            Assert.That(weekNumberProp!.PropertyType, Is.EqualTo(typeof(int)));

            Assert.That(referenteProp, Is.Not.Null);
            Assert.That(referenteProp!.PropertyType, Is.EqualTo(typeof(string)));
        }

        [Test]
        public void HeaderInfo_HasExactlyFourProperties()
        {
            // Arrange
            var properties = typeof(HeaderInfo).GetProperties();

            // Act & Assert - Verify the model has exactly 4 properties as specified
            Assert.That(properties.Length, Is.EqualTo(4));
        }
    }
}
