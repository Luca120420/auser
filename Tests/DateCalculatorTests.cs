using System;
using NUnit.Framework;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for the DateCalculator class.
    /// Validates: Requirements 5.3, 5.4, 5.5
    /// </summary>
    [TestFixture]
    public class DateCalculatorTests
    {
        private IDateCalculator _dateCalculator = null!;

        [SetUp]
        public void SetUp()
        {
            _dateCalculator = new DateCalculator();
        }

        #region AddDays Tests

        [Test]
        public void AddDays_AddSevenDays_ReturnsCorrectDate()
        {
            // Arrange
            var startDate = new DateTime(2024, 1, 26);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 2, 2)));
        }

        [Test]
        public void AddDays_AddSevenDaysAcrossMonthBoundary_ReturnsCorrectDate()
        {
            // Arrange
            var startDate = new DateTime(2024, 1, 29);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 2, 5)));
        }

        [Test]
        public void AddDays_AddSevenDaysAcrossYearBoundary_ReturnsCorrectDate()
        {
            // Arrange
            var startDate = new DateTime(2023, 12, 29);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 1, 5)));
        }

        [Test]
        public void AddDays_AddNegativeDays_ReturnsCorrectDate()
        {
            // Arrange
            var startDate = new DateTime(2024, 2, 5);

            // Act
            var result = _dateCalculator.AddDays(startDate, -7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 1, 29)));
        }

        [Test]
        public void AddDays_AddZeroDays_ReturnsSameDate()
        {
            // Arrange
            var startDate = new DateTime(2024, 1, 26);

            // Act
            var result = _dateCalculator.AddDays(startDate, 0);

            // Assert
            Assert.That(result, Is.EqualTo(startDate));
        }

        [Test]
        public void AddDays_LeapYearFebruary_HandlesCorrectly()
        {
            // Arrange - 2024 is a leap year
            var startDate = new DateTime(2024, 2, 26);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 3, 4)));
        }

        #endregion

        #region GetItalianMonthAbbreviation Tests

        [Test]
        [TestCase(1, "gen")]
        [TestCase(2, "feb")]
        [TestCase(3, "mar")]
        [TestCase(4, "apr")]
        [TestCase(5, "mag")]
        [TestCase(6, "giu")]
        [TestCase(7, "lug")]
        [TestCase(8, "ago")]
        [TestCase(9, "set")]
        [TestCase(10, "ott")]
        [TestCase(11, "nov")]
        [TestCase(12, "dic")]
        public void GetItalianMonthAbbreviation_ValidMonth_ReturnsCorrectAbbreviation(int month, string expected)
        {
            // Act
            var result = _dateCalculator.GetItalianMonthAbbreviation(month);

            // Assert
            Assert.That(result, Is.EqualTo(expected));
        }

        [Test]
        public void GetItalianMonthAbbreviation_MonthZero_ThrowsArgumentOutOfRangeException()
        {
            // Act & Assert
            var ex = Assert.Throws<ArgumentOutOfRangeException>(() => 
                _dateCalculator.GetItalianMonthAbbreviation(0));
            Assert.That(ex!.ParamName, Is.EqualTo("month"));
        }

        [Test]
        public void GetItalianMonthAbbreviation_MonthThirteen_ThrowsArgumentOutOfRangeException()
        {
            // Act & Assert
            var ex = Assert.Throws<ArgumentOutOfRangeException>(() => 
                _dateCalculator.GetItalianMonthAbbreviation(13));
            Assert.That(ex!.ParamName, Is.EqualTo("month"));
        }

        [Test]
        public void GetItalianMonthAbbreviation_NegativeMonth_ThrowsArgumentOutOfRangeException()
        {
            // Act & Assert
            var ex = Assert.Throws<ArgumentOutOfRangeException>(() => 
                _dateCalculator.GetItalianMonthAbbreviation(-1));
            Assert.That(ex!.ParamName, Is.EqualTo("month"));
        }

        #endregion

        #region FormatItalianDate Tests

        [Test]
        public void FormatItalianDate_January26_ReturnsCorrectFormat()
        {
            // Arrange
            var date = new DateTime(2024, 1, 26);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("26 gen"));
        }

        [Test]
        public void FormatItalianDate_February1_ReturnsCorrectFormat()
        {
            // Arrange
            var date = new DateTime(2024, 2, 1);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("01 feb"));
        }

        [Test]
        public void FormatItalianDate_December31_ReturnsCorrectFormat()
        {
            // Arrange
            var date = new DateTime(2024, 12, 31);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("31 dic"));
        }

        [Test]
        public void FormatItalianDate_SingleDigitDay_PadsWithZero()
        {
            // Arrange
            var date = new DateTime(2024, 5, 5);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("05 mag"));
        }

        [Test]
        public void FormatItalianDate_AllMonths_ReturnsCorrectAbbreviations()
        {
            // Arrange & Act & Assert
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 1, 15)), Is.EqualTo("15 gen"));
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 2, 15)), Is.EqualTo("15 feb"));
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 3, 15)), Is.EqualTo("15 mar"));
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 4, 15)), Is.EqualTo("15 apr"));
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 5, 15)), Is.EqualTo("15 mag"));
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 6, 15)), Is.EqualTo("15 giu"));
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 7, 15)), Is.EqualTo("15 lug"));
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 8, 15)), Is.EqualTo("15 ago"));
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 9, 15)), Is.EqualTo("15 set"));
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 10, 15)), Is.EqualTo("15 ott"));
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 11, 15)), Is.EqualTo("15 nov"));
            Assert.That(_dateCalculator.FormatItalianDate(new DateTime(2024, 12, 15)), Is.EqualTo("15 dic"));
        }

        #endregion

        #region ParseItalianDate Tests

        [Test]
        public void ParseItalianDate_ValidDate_ReturnsCorrectDateTime()
        {
            // Arrange
            var dateText = "26 gen";
            var year = 2024;

            // Act
            var result = _dateCalculator.ParseItalianDate(dateText, year);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 1, 26)));
        }

        [Test]
        public void ParseItalianDate_SingleDigitDay_ParsesCorrectly()
        {
            // Arrange
            var dateText = "1 feb";
            var year = 2024;

            // Act
            var result = _dateCalculator.ParseItalianDate(dateText, year);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 2, 1)));
        }

        [Test]
        public void ParseItalianDate_TwoDigitDay_ParsesCorrectly()
        {
            // Arrange
            var dateText = "01 feb";
            var year = 2024;

            // Act
            var result = _dateCalculator.ParseItalianDate(dateText, year);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 2, 1)));
        }

        [Test]
        public void ParseItalianDate_AllMonths_ParsesCorrectly()
        {
            // Arrange & Act & Assert
            Assert.That(_dateCalculator.ParseItalianDate("15 gen", 2024), Is.EqualTo(new DateTime(2024, 1, 15)));
            Assert.That(_dateCalculator.ParseItalianDate("15 feb", 2024), Is.EqualTo(new DateTime(2024, 2, 15)));
            Assert.That(_dateCalculator.ParseItalianDate("15 mar", 2024), Is.EqualTo(new DateTime(2024, 3, 15)));
            Assert.That(_dateCalculator.ParseItalianDate("15 apr", 2024), Is.EqualTo(new DateTime(2024, 4, 15)));
            Assert.That(_dateCalculator.ParseItalianDate("15 mag", 2024), Is.EqualTo(new DateTime(2024, 5, 15)));
            Assert.That(_dateCalculator.ParseItalianDate("15 giu", 2024), Is.EqualTo(new DateTime(2024, 6, 15)));
            Assert.That(_dateCalculator.ParseItalianDate("15 lug", 2024), Is.EqualTo(new DateTime(2024, 7, 15)));
            Assert.That(_dateCalculator.ParseItalianDate("15 ago", 2024), Is.EqualTo(new DateTime(2024, 8, 15)));
            Assert.That(_dateCalculator.ParseItalianDate("15 set", 2024), Is.EqualTo(new DateTime(2024, 9, 15)));
            Assert.That(_dateCalculator.ParseItalianDate("15 ott", 2024), Is.EqualTo(new DateTime(2024, 10, 15)));
            Assert.That(_dateCalculator.ParseItalianDate("15 nov", 2024), Is.EqualTo(new DateTime(2024, 11, 15)));
            Assert.That(_dateCalculator.ParseItalianDate("15 dic", 2024), Is.EqualTo(new DateTime(2024, 12, 15)));
        }

        [Test]
        public void ParseItalianDate_CaseInsensitive_ParsesCorrectly()
        {
            // Arrange & Act & Assert
            Assert.That(_dateCalculator.ParseItalianDate("26 GEN", 2024), Is.EqualTo(new DateTime(2024, 1, 26)));
            Assert.That(_dateCalculator.ParseItalianDate("26 Gen", 2024), Is.EqualTo(new DateTime(2024, 1, 26)));
            Assert.That(_dateCalculator.ParseItalianDate("26 gEn", 2024), Is.EqualTo(new DateTime(2024, 1, 26)));
        }

        [Test]
        public void ParseItalianDate_ExtraSpaces_ParsesCorrectly()
        {
            // Arrange
            var dateText = "  26   gen  ";
            var year = 2024;

            // Act
            var result = _dateCalculator.ParseItalianDate(dateText, year);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 1, 26)));
        }

        [Test]
        public void ParseItalianDate_EmptyString_ThrowsFormatException()
        {
            // Act & Assert
            var ex = Assert.Throws<FormatException>(() => 
                _dateCalculator.ParseItalianDate("", 2024));
            Assert.That(ex!.Message, Does.Contain("vuoto"));
        }

        [Test]
        public void ParseItalianDate_NullString_ThrowsFormatException()
        {
            // Act & Assert
            var ex = Assert.Throws<FormatException>(() => 
                _dateCalculator.ParseItalianDate(null!, 2024));
            Assert.That(ex!.Message, Does.Contain("vuoto"));
        }

        [Test]
        public void ParseItalianDate_InvalidFormat_ThrowsFormatException()
        {
            // Act & Assert
            var ex = Assert.Throws<FormatException>(() => 
                _dateCalculator.ParseItalianDate("26-gen", 2024));
            Assert.That(ex!.Message, Does.Contain("Formato data non valido"));
        }

        [Test]
        public void ParseItalianDate_InvalidDay_ThrowsFormatException()
        {
            // Act & Assert
            var ex = Assert.Throws<FormatException>(() => 
                _dateCalculator.ParseItalianDate("abc gen", 2024));
            Assert.That(ex!.Message, Does.Contain("Giorno non valido"));
        }

        [Test]
        public void ParseItalianDate_DayOutOfRange_ThrowsFormatException()
        {
            // Act & Assert
            var ex = Assert.Throws<FormatException>(() => 
                _dateCalculator.ParseItalianDate("32 gen", 2024));
            Assert.That(ex!.Message, Does.Contain("Giorno non valido"));
        }

        [Test]
        public void ParseItalianDate_InvalidMonth_ThrowsFormatException()
        {
            // Act & Assert
            var ex = Assert.Throws<FormatException>(() => 
                _dateCalculator.ParseItalianDate("26 xyz", 2024));
            Assert.That(ex!.Message, Does.Contain("Abbreviazione mese non valida"));
        }

        [Test]
        public void ParseItalianDate_InvalidDayForMonth_ThrowsFormatException()
        {
            // Act & Assert - February 30 doesn't exist
            var ex = Assert.Throws<FormatException>(() => 
                _dateCalculator.ParseItalianDate("30 feb", 2024));
            Assert.That(ex!.Message, Does.Contain("Data non valida"));
        }

        [Test]
        public void ParseItalianDate_LeapYearFebruary29_ParsesCorrectly()
        {
            // Arrange - 2024 is a leap year
            var dateText = "29 feb";
            var year = 2024;

            // Act
            var result = _dateCalculator.ParseItalianDate(dateText, year);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 2, 29)));
        }

        [Test]
        public void ParseItalianDate_NonLeapYearFebruary29_ThrowsFormatException()
        {
            // Act & Assert - 2023 is not a leap year
            var ex = Assert.Throws<FormatException>(() => 
                _dateCalculator.ParseItalianDate("29 feb", 2023));
            Assert.That(ex!.Message, Does.Contain("Data non valida"));
        }

        #endregion

        #region Edge Case Tests - Year Boundary Transitions

        [Test]
        public void AddDays_December31ToJanuary7_CrossesYearBoundaryCorrectly()
        {
            // Arrange - Dec 31, 2023 (Sunday) + 7 days = Jan 7, 2024 (Sunday)
            var startDate = new DateTime(2023, 12, 31);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 1, 7)));
            Assert.That(result.Year, Is.EqualTo(2024));
            Assert.That(result.Month, Is.EqualTo(1));
            Assert.That(result.Day, Is.EqualTo(7));
        }

        [Test]
        public void FormatItalianDate_December31_ReturnsCorrectFormatBeforeYearBoundary()
        {
            // Arrange
            var date = new DateTime(2023, 12, 31);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("31 dic"));
        }

        [Test]
        public void FormatItalianDate_January1_ReturnsCorrectFormatAfterYearBoundary()
        {
            // Arrange
            var date = new DateTime(2024, 1, 1);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("01 gen"));
        }

        [Test]
        public void FormatItalianDate_January7_ReturnsCorrectFormatAfterYearBoundary()
        {
            // Arrange
            var date = new DateTime(2024, 1, 7);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("07 gen"));
        }

        [Test]
        public void ParseItalianDate_December31_ParsesCorrectlyBeforeYearBoundary()
        {
            // Arrange
            var dateText = "31 dic";
            var year = 2023;

            // Act
            var result = _dateCalculator.ParseItalianDate(dateText, year);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2023, 12, 31)));
        }

        [Test]
        public void ParseItalianDate_January7_ParsesCorrectlyAfterYearBoundary()
        {
            // Arrange
            var dateText = "07 gen";
            var year = 2024;

            // Act
            var result = _dateCalculator.ParseItalianDate(dateText, year);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 1, 7)));
        }

        [Test]
        public void AddDaysAndFormat_December31ToJanuary7_FormatsCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2023, 12, 31);

            // Act
            var newDate = _dateCalculator.AddDays(startDate, 7);
            var formatted = _dateCalculator.FormatItalianDate(newDate);

            // Assert
            Assert.That(formatted, Is.EqualTo("07 gen"));
        }

        #endregion

        #region Edge Case Tests - Month Boundary Transitions

        [Test]
        public void AddDays_January29ToFebruary5_CrossesMonthBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2024, 1, 29);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 2, 5)));
        }

        [Test]
        public void AddDays_February26ToMarch4_CrossesMonthBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2024, 2, 26);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 3, 4)));
        }

        [Test]
        public void AddDays_March25ToApril1_CrossesMonthBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2024, 3, 25);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 4, 1)));
        }

        [Test]
        public void AddDays_April29ToMay6_CrossesMonthBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2024, 4, 29);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 5, 6)));
        }

        [Test]
        public void AddDays_May27ToJune3_CrossesMonthBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2024, 5, 27);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 6, 3)));
        }

        [Test]
        public void AddDays_June24ToJuly1_CrossesMonthBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2024, 6, 24);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 7, 1)));
        }

        [Test]
        public void AddDays_July29ToAugust5_CrossesMonthBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2024, 7, 29);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 8, 5)));
        }

        [Test]
        public void AddDays_August26ToSeptember2_CrossesMonthBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2024, 8, 26);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 9, 2)));
        }

        [Test]
        public void AddDays_September30ToOctober7_CrossesMonthBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2024, 9, 30);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 10, 7)));
        }

        [Test]
        public void AddDays_October28ToNovember4_CrossesMonthBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2024, 10, 28);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 11, 4)));
        }

        [Test]
        public void AddDays_November25ToDecember2_CrossesMonthBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2024, 11, 25);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 12, 2)));
        }

        [Test]
        public void AddDays_December30ToJanuary6_CrossesMonthAndYearBoundaryCorrectly()
        {
            // Arrange
            var startDate = new DateTime(2023, 12, 30);

            // Act
            var result = _dateCalculator.AddDays(startDate, 7);

            // Assert
            Assert.That(result, Is.EqualTo(new DateTime(2024, 1, 6)));
        }

        #endregion

        #region Edge Case Tests - All Month Abbreviations

        [Test]
        public void FormatItalianDate_January_ReturnsGenAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 1, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 gen"));
        }

        [Test]
        public void FormatItalianDate_February_ReturnsFebAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 2, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 feb"));
        }

        [Test]
        public void FormatItalianDate_March_ReturnsMarAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 3, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 mar"));
        }

        [Test]
        public void FormatItalianDate_April_ReturnsAprAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 4, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 apr"));
        }

        [Test]
        public void FormatItalianDate_May_ReturnsMagAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 5, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 mag"));
        }

        [Test]
        public void FormatItalianDate_June_ReturnsGiuAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 6, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 giu"));
        }

        [Test]
        public void FormatItalianDate_July_ReturnsLugAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 7, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 lug"));
        }

        [Test]
        public void FormatItalianDate_August_ReturnsAgoAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 8, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 ago"));
        }

        [Test]
        public void FormatItalianDate_September_ReturnsSetAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 9, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 set"));
        }

        [Test]
        public void FormatItalianDate_October_ReturnsOttAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 10, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 ott"));
        }

        [Test]
        public void FormatItalianDate_November_ReturnsNovAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 11, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 nov"));
        }

        [Test]
        public void FormatItalianDate_December_ReturnsDicAbbreviation()
        {
            // Arrange
            var date = new DateTime(2024, 12, 15);

            // Act
            var result = _dateCalculator.FormatItalianDate(date);

            // Assert
            Assert.That(result, Is.EqualTo("15 dic"));
        }

        [Test]
        public void ParseItalianDate_AllMonthAbbreviations_ParseCorrectly()
        {
            // Arrange & Act & Assert - Testing all 12 month abbreviations
            Assert.That(_dateCalculator.ParseItalianDate("01 gen", 2024), Is.EqualTo(new DateTime(2024, 1, 1)));
            Assert.That(_dateCalculator.ParseItalianDate("01 feb", 2024), Is.EqualTo(new DateTime(2024, 2, 1)));
            Assert.That(_dateCalculator.ParseItalianDate("01 mar", 2024), Is.EqualTo(new DateTime(2024, 3, 1)));
            Assert.That(_dateCalculator.ParseItalianDate("01 apr", 2024), Is.EqualTo(new DateTime(2024, 4, 1)));
            Assert.That(_dateCalculator.ParseItalianDate("01 mag", 2024), Is.EqualTo(new DateTime(2024, 5, 1)));
            Assert.That(_dateCalculator.ParseItalianDate("01 giu", 2024), Is.EqualTo(new DateTime(2024, 6, 1)));
            Assert.That(_dateCalculator.ParseItalianDate("01 lug", 2024), Is.EqualTo(new DateTime(2024, 7, 1)));
            Assert.That(_dateCalculator.ParseItalianDate("01 ago", 2024), Is.EqualTo(new DateTime(2024, 8, 1)));
            Assert.That(_dateCalculator.ParseItalianDate("01 set", 2024), Is.EqualTo(new DateTime(2024, 9, 1)));
            Assert.That(_dateCalculator.ParseItalianDate("01 ott", 2024), Is.EqualTo(new DateTime(2024, 10, 1)));
            Assert.That(_dateCalculator.ParseItalianDate("01 nov", 2024), Is.EqualTo(new DateTime(2024, 11, 1)));
            Assert.That(_dateCalculator.ParseItalianDate("01 dic", 2024), Is.EqualTo(new DateTime(2024, 12, 1)));
        }

        [Test]
        public void GetItalianMonthAbbreviation_AllTwelveMonths_ReturnsCorrectAbbreviations()
        {
            // Act & Assert - Comprehensive test for all 12 months
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(1), Is.EqualTo("gen"));
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(2), Is.EqualTo("feb"));
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(3), Is.EqualTo("mar"));
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(4), Is.EqualTo("apr"));
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(5), Is.EqualTo("mag"));
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(6), Is.EqualTo("giu"));
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(7), Is.EqualTo("lug"));
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(8), Is.EqualTo("ago"));
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(9), Is.EqualTo("set"));
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(10), Is.EqualTo("ott"));
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(11), Is.EqualTo("nov"));
            Assert.That(_dateCalculator.GetItalianMonthAbbreviation(12), Is.EqualTo("dic"));
        }

        #endregion

        #region Integration Tests

        [Test]
        public void FormatAndParse_RoundTrip_ReturnsOriginalDate()
        {
            // Arrange
            var originalDate = new DateTime(2024, 1, 26);

            // Act
            var formatted = _dateCalculator.FormatItalianDate(originalDate);
            var parsed = _dateCalculator.ParseItalianDate(formatted, originalDate.Year);

            // Assert
            Assert.That(parsed, Is.EqualTo(originalDate));
        }

        [Test]
        public void AddDaysAndFormat_WeekTransition_ReturnsCorrectFormattedDate()
        {
            // Arrange
            var startDate = new DateTime(2024, 1, 26);

            // Act
            var newDate = _dateCalculator.AddDays(startDate, 7);
            var formatted = _dateCalculator.FormatItalianDate(newDate);

            // Assert
            Assert.That(formatted, Is.EqualTo("02 feb"));
        }

        #endregion
    }
}
