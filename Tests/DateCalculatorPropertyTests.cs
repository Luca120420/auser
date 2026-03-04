using System;
using FsCheck;
using NUnit.Framework;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for DateCalculator class using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// Validates: Requirements 5.3, 5.4, 5.5
    /// </summary>
    [TestFixture]
    public class DateCalculatorPropertyTests
    {
        private IDateCalculator _dateCalculator = null!;

        [SetUp]
        public void SetUp()
        {
            _dateCalculator = new DateCalculator();
        }

        // Feature: auser-excel-transformer, Property 17: Week Date Arithmetic
        /// <summary>
        /// Property 17: Week Date Arithmetic
        /// For any date, adding 7 days should produce a date exactly one week later,
        /// and this should work correctly across month and year boundaries.
        /// **Validates: Requirements 5.3, 5.4**
        /// </summary>
        [Test]
        public void Property_WeekDateArithmetic()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryDate(),
                (DateTime date) =>
                {
                    // Act - Add 7 days
                    var result = _dateCalculator.AddDays(date, 7);

                    // Assert - Result should be exactly 7 days later
                    var expectedDate = date.AddDays(7);
                    var datesMatch = result == expectedDate;

                    // Assert - The difference should be exactly 7 days
                    var daysDifference = (result - date).TotalDays;
                    var differenceIsSevenDays = Math.Abs(daysDifference - 7.0) < 0.0001;

                    // Assert - Day of week should be the same
                    var dayOfWeekMatches = result.DayOfWeek == date.DayOfWeek;

                    return datesMatch && differenceIsSevenDays && dayOfWeekMatches;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 17: Week Date Arithmetic (Month Boundaries)
        /// <summary>
        /// Property 17: Week Date Arithmetic - Month Boundary Test
        /// For any date near the end of a month, adding 7 days should correctly
        /// handle the month transition when it occurs.
        /// **Validates: Requirements 5.3, 5.4**
        /// </summary>
        [Test]
        public void Property_WeekDateArithmetic_MonthBoundaries()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryDateNearMonthEnd(),
                (DateTime date) =>
                {
                    // Act - Add 7 days
                    var result = _dateCalculator.AddDays(date, 7);

                    // Assert - Result should be exactly 7 days later
                    var expectedDate = date.AddDays(7);
                    var datesMatch = result == expectedDate;

                    // Assert - The difference should be exactly 7 days
                    var daysDifference = (result - date).TotalDays;
                    var differenceIsSevenDays = Math.Abs(daysDifference - 7.0) < 0.0001;

                    // Assert - Day of week should be the same
                    var dayOfWeekMatches = result.DayOfWeek == date.DayOfWeek;
                    
                    return datesMatch && differenceIsSevenDays && dayOfWeekMatches;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 17: Week Date Arithmetic (Year Boundaries)
        /// <summary>
        /// Property 17: Week Date Arithmetic - Year Boundary Test
        /// For any date near the end of a year, adding 7 days should correctly
        /// handle the year transition.
        /// **Validates: Requirements 5.3, 5.4**
        /// </summary>
        [Test]
        public void Property_WeekDateArithmetic_YearBoundaries()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryDateNearYearEnd(),
                (DateTime date) =>
                {
                    // Act - Add 7 days
                    var result = _dateCalculator.AddDays(date, 7);

                    // Assert - Result should be exactly 7 days later
                    var expectedDate = date.AddDays(7);
                    var datesMatch = result == expectedDate;

                    // Assert - The difference should be exactly 7 days
                    var daysDifference = (result - date).TotalDays;
                    var differenceIsSevenDays = Math.Abs(daysDifference - 7.0) < 0.0001;

                    // Assert - We should have crossed into the next year
                    var crossedYearBoundary = result.Year == date.Year + 1;
                    
                    return datesMatch && differenceIsSevenDays && crossedYearBoundary;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 18: Italian Month Abbreviation Mapping
        /// <summary>
        /// Property 18: Italian Month Abbreviation Mapping
        /// For any month number (1-12), the application should produce the correct Italian month abbreviation:
        /// 1→gen, 2→feb, 3→mar, 4→apr, 5→mag, 6→giu, 7→lug, 8→ago, 9→set, 10→ott, 11→nov, 12→dic.
        /// **Validates: Requirements 5.5**
        /// </summary>
        [Test]
        public void Property_ItalianMonthAbbreviationMapping()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Define the expected mapping
            var expectedMapping = new System.Collections.Generic.Dictionary<int, string>
            {
                { 1, "gen" },
                { 2, "feb" },
                { 3, "mar" },
                { 4, "apr" },
                { 5, "mag" },
                { 6, "giu" },
                { 7, "lug" },
                { 8, "ago" },
                { 9, "set" },
                { 10, "ott" },
                { 11, "nov" },
                { 12, "dic" }
            };

            Prop.ForAll(
                Arb.From(Gen.Choose(1, 12)),
                (int month) =>
                {
                    // Act
                    var result = _dateCalculator.GetItalianMonthAbbreviation(month);

                    // Assert - Result should match the expected abbreviation
                    var expectedAbbreviation = expectedMapping[month];
                    var abbreviationMatches = result == expectedAbbreviation;

                    // Assert - Result should be exactly 3 characters long
                    var lengthIsThree = result.Length == 3;

                    // Assert - Result should be lowercase
                    var isLowercase = result == result.ToLower();

                    return abbreviationMatches && lengthIsThree && isLowercase;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 18: Italian Month Abbreviation Mapping (Invalid Months)
        /// <summary>
        /// Property 18: Italian Month Abbreviation Mapping - Invalid Month Test
        /// For any month number outside the range 1-12, the application should throw
        /// an ArgumentOutOfRangeException.
        /// **Validates: Requirements 5.5**
        /// </summary>
        [Test]
        public void Property_ItalianMonthAbbreviationMapping_InvalidMonths()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(Gen.Choose(-100, 100).Where(m => m < 1 || m > 12)),
                (int month) =>
                {
                    // Act & Assert - Should throw ArgumentOutOfRangeException
                    try
                    {
                        _dateCalculator.GetItalianMonthAbbreviation(month);
                        return false; // Should have thrown an exception
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        return true; // Expected exception
                    }
                    catch
                    {
                        return false; // Wrong exception type
                    }
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 17 & 18: Date Formatting Round-Trip
        /// <summary>
        /// Property: Date Formatting Round-Trip
        /// For any date, formatting it to Italian format and then parsing it back
        /// should produce the original date.
        /// **Validates: Requirements 5.3, 5.4, 5.5**
        /// </summary>
        [Test]
        public void Property_DateFormattingRoundTrip()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryDate(),
                (DateTime date) =>
                {
                    // Act - Format and then parse
                    var formatted = _dateCalculator.FormatItalianDate(date);
                    var parsed = _dateCalculator.ParseItalianDate(formatted, date.Year);

                    // Assert - Parsed date should equal original date
                    return parsed == date;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 17: Week Addition Commutativity
        /// <summary>
        /// Property: Week Addition Commutativity
        /// Adding 7 days twice should be the same as adding 14 days once.
        /// **Validates: Requirements 5.3, 5.4**
        /// </summary>
        [Test]
        public void Property_WeekAdditionCommutativity()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryDate(),
                (DateTime date) =>
                {
                    // Act - Add 7 days twice
                    var result1 = _dateCalculator.AddDays(date, 7);
                    var result2 = _dateCalculator.AddDays(result1, 7);

                    // Act - Add 14 days once
                    var result3 = _dateCalculator.AddDays(date, 14);

                    // Assert - Both results should be the same
                    return result2 == result3;
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 17: Week Addition Inverse
        /// <summary>
        /// Property: Week Addition Inverse
        /// Adding 7 days and then subtracting 7 days should return the original date.
        /// **Validates: Requirements 5.3, 5.4**
        /// </summary>
        [Test]
        public void Property_WeekAdditionInverse()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                ArbitraryDate(),
                (DateTime date) =>
                {
                    // Act - Add 7 days and then subtract 7 days
                    var result1 = _dateCalculator.AddDays(date, 7);
                    var result2 = _dateCalculator.AddDays(result1, -7);

                    // Assert - Should return to original date
                    return result2 == date;
                }
            ).Check(config);
        }

        #region Custom Generators

        /// <summary>
        /// Generator for arbitrary dates between 2020 and 2030
        /// </summary>
        private static Arbitrary<DateTime> ArbitraryDate()
        {
            var dateGen = from year in Gen.Choose(2020, 2030)
                         from month in Gen.Choose(1, 12)
                         from day in Gen.Choose(1, 28) // Use 28 to avoid invalid dates
                         select new DateTime(year, month, day);

            return Arb.From(dateGen);
        }

        /// <summary>
        /// Generator for dates near the end of a month (last 7 days)
        /// </summary>
        private static Arbitrary<DateTime> ArbitraryDateNearMonthEnd()
        {
            var dateGen = from year in Gen.Choose(2020, 2030)
                         from month in Gen.Choose(1, 12)
                         let daysInMonth = DateTime.DaysInMonth(year, month)
                         from day in Gen.Choose(Math.Max(1, daysInMonth - 7), daysInMonth)
                         select new DateTime(year, month, day);

            return Arb.From(dateGen);
        }

        /// <summary>
        /// Generator for dates near the end of a year (last 7 days of December)
        /// </summary>
        private static Arbitrary<DateTime> ArbitraryDateNearYearEnd()
        {
            var dateGen = from year in Gen.Choose(2020, 2029) // Avoid 2030 to ensure we can add 7 days
                         from day in Gen.Choose(25, 31)
                         select new DateTime(year, 12, day);

            return Arb.From(dateGen);
        }

        #endregion
    }
}
