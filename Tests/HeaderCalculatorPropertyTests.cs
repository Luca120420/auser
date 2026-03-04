using System;
using System.Text.RegularExpressions;
using FsCheck;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for HeaderCalculator class using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// Validates: Requirements 5.1, 5.2, 5.6, 5.7, 5.8
    /// </summary>
    [TestFixture]
    public class HeaderCalculatorPropertyTests
    {
        private IHeaderCalculator _headerCalculator = null!;
        private IDateCalculator _dateCalculator = null!;

        [SetUp]
        public void SetUp()
        {
            _dateCalculator = new DateCalculator();
            _headerCalculator = new HeaderCalculator(_dateCalculator);
        }

        // Feature: auser-excel-transformer, Property 16: Header Date Parsing
        /// <summary>
        /// Property 16: Header Date Parsing
        /// For any valid header string in the format "DD mmm DD mmm Settimana N...",
        /// the application should successfully extract both dates and the week number.
        /// **Validates: Requirements 5.1, 5.2**
        /// </summary>
        [Test]
        public void Property_HeaderDateParsing()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryValidHeader(),
                (ValidHeaderData headerData) =>
                {
                    // Arrange - Create a valid header string
                    var headerText = headerData.ToHeaderString();

                    try
                    {
                        // Act - Parse the header
                        var result = _headerCalculator.ParseHeader(headerText);

                        // Assert - Monday date should be extracted correctly
                        var mondayMatches = result.MondayDate.Day == headerData.MondayDay &&
                                          result.MondayDate.Month == headerData.MondayMonth;

                        // Assert - Sunday date should be extracted correctly
                        var sundayMatches = result.SundayDate.Day == headerData.SundayDay &&
                                          result.SundayDate.Month == headerData.SundayMonth;

                        // Assert - Week number should be extracted correctly
                        var weekNumberMatches = result.WeekNumber == headerData.WeekNumber;

                        // Assert - Referente should be extracted correctly
                        var referenteMatches = result.Referente == headerData.Referente;

                        // Assert - Sunday should be after Monday (or same week)
                        var sundayAfterMonday = result.SundayDate >= result.MondayDate;

                        // Assert - The difference should be at most 7 days
                        var daysDifference = (result.SundayDate - result.MondayDate).TotalDays;
                        var differenceValid = daysDifference >= 0 && daysDifference <= 7;

                        return mondayMatches && sundayMatches && weekNumberMatches && 
                               referenteMatches && sundayAfterMonday && differenceValid;
                    }
                    catch (FormatException)
                    {
                        // If parsing fails, the property doesn't hold
                        return false;
                    }
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 19: Week Number Increment
        /// <summary>
        /// Property 19: Week Number Increment
        /// For any week number N in the previous header, the new header should contain week number N+1.
        /// **Validates: Requirements 5.6**
        /// </summary>
        [Test]
        public void Property_WeekNumberIncrement()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryValidHeader(),
                (ValidHeaderData headerData) =>
                {
                    // Arrange - Create a valid header string
                    var previousHeader = headerData.ToHeaderString();

                    try
                    {
                        // Act - Generate next week's header
                        var nextHeader = _headerCalculator.GenerateNextWeekHeader(previousHeader);

                        // Parse the new header to extract the week number
                        var parsedNext = _headerCalculator.ParseHeader(nextHeader);

                        // Assert - Week number should be incremented by 1
                        var weekNumberIncremented = parsedNext.WeekNumber == headerData.WeekNumber + 1;

                        // Assert - The new header should be parseable
                        var headerIsParseable = parsedNext != null;

                        return weekNumberIncremented && headerIsParseable;
                    }
                    catch (FormatException)
                    {
                        // If parsing or generation fails, the property doesn't hold
                        return false;
                    }
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 20: Referente Reset
        /// <summary>
        /// Property 20: Referente Reset
        /// For any generated header, the referente text should be exactly
        /// "Inserire nome e numero di telefono del referente".
        /// **Validates: Requirements 5.7**
        /// </summary>
        [Test]
        public void Property_ReferenteReset()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryValidHeader(),
                (ValidHeaderData headerData) =>
                {
                    // Arrange - Create a valid header string with any referente text
                    var previousHeader = headerData.ToHeaderString();

                    try
                    {
                        // Act - Generate next week's header
                        var nextHeader = _headerCalculator.GenerateNextWeekHeader(previousHeader);

                        // Parse the new header to extract the referente
                        var parsedNext = _headerCalculator.ParseHeader(nextHeader);

                        // Assert - Referente should be reset to the default text
                        var expectedReferente = "Inserire nome e numero di telefono del referente";
                        var referenteReset = parsedNext.Referente == expectedReferente;

                        // Assert - The new header should contain the referente text
                        var headerContainsReferente = nextHeader.Contains(expectedReferente);

                        // Assert - The new header should not contain the old referente text
                        // (unless it was already the default)
                        var oldReferenteRemoved = headerData.Referente == expectedReferente ||
                                                 !nextHeader.Contains(headerData.Referente);

                        return referenteReset && headerContainsReferente && oldReferenteRemoved;
                    }
                    catch (FormatException)
                    {
                        // If parsing or generation fails, the property doesn't hold
                        return false;
                    }
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 21: Header Format Compliance
        /// <summary>
        /// Property 21: Header Format Compliance
        /// For any generated header, it should match the format
        /// "DD mmm DD mmm Settimana Nreferente settimana = Inserire nome e numero di telefono del referente"
        /// where DD are day numbers, mmm are Italian month abbreviations, and N is the week number.
        /// **Validates: Requirements 5.8**
        /// </summary>
        [Test]
        public void Property_HeaderFormatCompliance()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryValidHeader(),
                (ValidHeaderData headerData) =>
                {
                    // Arrange - Create a valid header string
                    var previousHeader = headerData.ToHeaderString();

                    try
                    {
                        // Act - Generate next week's header
                        var nextHeader = _headerCalculator.GenerateNextWeekHeader(previousHeader);

                        // Assert - Header should match the expected format pattern
                        // Format: "DD mmm DD mmm Settimana Nreferente settimana = Inserire nome e numero di telefono del referente"
                        var formatPattern = @"^\d{2} \w{3} \d{2} \w{3} Settimana \d+referente settimana = .+$";
                        var matchesFormat = Regex.IsMatch(nextHeader, formatPattern);

                        // Assert - Header should contain "Settimana"
                        var containsSettimana = nextHeader.Contains("Settimana");

                        // Assert - Header should contain "referente settimana ="
                        var containsReferentePrefix = nextHeader.Contains("referente settimana =");

                        // Assert - Header should contain the default referente text
                        var containsDefaultReferente = nextHeader.Contains("Inserire nome e numero di telefono del referente");

                        // Assert - Header should be parseable (round-trip test)
                        var parsedHeader = _headerCalculator.ParseHeader(nextHeader);
                        var isParseable = parsedHeader != null;

                        // Assert - Parsed dates should have valid Italian month abbreviations
                        var validMonthAbbreviations = new[] { "gen", "feb", "mar", "apr", "mag", "giu", 
                                                              "lug", "ago", "set", "ott", "nov", "dic" };
                        var mondayMonth = _dateCalculator.GetItalianMonthAbbreviation(parsedHeader.MondayDate.Month);
                        var sundayMonth = _dateCalculator.GetItalianMonthAbbreviation(parsedHeader.SundayDate.Month);
                        var hasValidMonths = System.Array.IndexOf(validMonthAbbreviations, mondayMonth) >= 0 &&
                                           System.Array.IndexOf(validMonthAbbreviations, sundayMonth) >= 0;

                        // Assert - Header should contain the correct month abbreviations
                        var containsMondayMonth = nextHeader.Contains(mondayMonth);
                        var containsSundayMonth = nextHeader.Contains(sundayMonth);

                        return matchesFormat && containsSettimana && containsReferentePrefix &&
                               containsDefaultReferente && isParseable && hasValidMonths &&
                               containsMondayMonth && containsSundayMonth;
                    }
                    catch (FormatException)
                    {
                        // If parsing or generation fails, the property doesn't hold
                        return false;
                    }
                }
            ).Check(config);
        }

        #region Custom Generators

        /// <summary>
        /// Data class for generating valid header strings
        /// </summary>
        public class ValidHeaderData
        {
            public int MondayDay { get; set; }
            public int MondayMonth { get; set; }
            public int SundayDay { get; set; }
            public int SundayMonth { get; set; }
            public int WeekNumber { get; set; }
            public string Referente { get; set; } = "";

            public string ToHeaderString()
            {
                var italianMonths = new[] { "gen", "feb", "mar", "apr", "mag", "giu", 
                                           "lug", "ago", "set", "ott", "nov", "dic" };
                
                var mondayMonthAbbr = italianMonths[MondayMonth - 1];
                var sundayMonthAbbr = italianMonths[SundayMonth - 1];

                return $"{MondayDay:D2} {mondayMonthAbbr} {SundayDay:D2} {sundayMonthAbbr} Settimana {WeekNumber}referente settimana = {Referente}";
            }
        }

        /// <summary>
        /// Generator for valid header data
        /// </summary>
        private static Arbitrary<ValidHeaderData> ArbitraryValidHeader()
        {
            var italianNames = new[] { "Mario Rossi", "Luigi Bianchi", "Giuseppe Verdi", "Anna Ferrari", 
                                      "Maria Colombo", "Francesco Romano", "Test User", 
                                      "Inserire nome e numero di telefono del referente" };
            
            var phoneNumbers = new[] { "333-1234567", "340-9876543", "347-5551234", "320-1112233", "" };

            var headerGen = from weekNumber in Gen.Choose(1, 52)
                           from mondayMonth in Gen.Choose(1, 12)
                           from mondayDay in Gen.Choose(1, 28) // Use 28 to avoid invalid dates
                           from name in Gen.Elements(italianNames)
                           from phone in Gen.Elements(phoneNumbers)
                           let referente = string.IsNullOrEmpty(phone) ? name : $"{name} {phone}"
                           let year = DateTime.Now.Year
                           let mondayDate = new DateTime(year, mondayMonth, mondayDay)
                           let sundayDate = mondayDate.AddDays(6) // Sunday is 6 days after Monday
                           select new ValidHeaderData
                           {
                               MondayDay = mondayDay,
                               MondayMonth = mondayMonth,
                               SundayDay = sundayDate.Day,
                               SundayMonth = sundayDate.Month,
                               WeekNumber = weekNumber,
                               Referente = referente
                           };

            return Arb.From(headerGen);
        }

        /// <summary>
        /// Generator for headers near year boundaries
        /// </summary>
        private static Arbitrary<ValidHeaderData> ArbitraryHeaderNearYearBoundary()
        {
            var headerGen = from weekNumber in Gen.Choose(51, 53)
                           from mondayDay in Gen.Choose(25, 31)
                           let year = DateTime.Now.Year
                           let mondayDate = new DateTime(year, 12, mondayDay)
                           let sundayDate = mondayDate.AddDays(6)
                           select new ValidHeaderData
                           {
                               MondayDay = mondayDay,
                               MondayMonth = 12,
                               SundayDay = sundayDate.Day,
                               SundayMonth = sundayDate.Month,
                               WeekNumber = weekNumber,
                               Referente = "Test User"
                           };

            return Arb.From(headerGen);
        }

        #endregion
    }
}
