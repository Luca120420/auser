# Test Data Files for Auser Excel Transformer

This directory contains sample test data files for testing the Auser Excel Transformer application. These files cover normal operations, edge cases, and error scenarios.

## Overview

The test data files are designed to validate:
- **Requirements 2.1**: CSV data parsing with all required columns
- **Requirements 3.1**: Excel workbook with multiple numbered sheets
- **Requirements 3.4**: Excel workbook with "fissi" sheet containing recurring appointments

## Test Files

### 1. sample_input.csv

**Purpose**: Standard test case with realistic Italian data

**Contents**:
- 14 service appointments with realistic Italian names, addresses, and dates
- Mix of "Accompag. con macchina attrezzata" and "Accomp. servizi con trasporto" activities
- Various statuses: PIANIFICATO, IN PIANIFICAZIONE, ANNULLATO
- Italian special characters (accented letters)
- Some appointments with DESCRIZIONE PUNTO PARTENZA, some without
- Some appointments with INDIRIZZO DESTINAZIONE, some without (to test conditional logic)

**Expected Behavior**:
- Should parse successfully
- Appointments with "ANNULLATO" status should be filtered out (3 appointments)
- Appointments with "Accompag. con macchina attrezzata" should be marked for yellow highlighting (4 appointments)
- DESCRIZIONE PUNTO PARTENZA should be duplicated when present
- INDIRIZZO column should use CAUSALE DESTINAZIONE when INDIRIZZO DESTINAZIONE is empty

**Test Coverage**:
- Normal CSV parsing (Req 2.1, 2.2)
- Italian character preservation (Req 2.4)
- Cancelled appointment filtering (Req 4.3)
- Yellow highlighting rule (Req 4.1)
- Conditional address concatenation (Req 4.6, 4.7)

---

### 2. edge_cases_input.csv

**Purpose**: Test edge cases and special scenarios

**Contents**:
- **Row 1**: DESCRIZIONE PUNTO PARTENZA with text to be duplicated
- **Row 2**: ANNULLATO status - should be filtered out
- **Row 3**: Empty INDIRIZZO DESTINAZIONE - should use CAUSALE DESTINAZIONE
- **Row 4**: ANNULLATO with "Accompag. con macchina attrezzata" - should be filtered (not highlighted)
- **Row 5**: Italian special characters (àèéìòù) in notes
- **Row 6**: Apostrophes and accents in names (O'Brien, Seán)
- **Row 7**: ANNULLATO with special characters - should be filtered
- **Row 8**: Empty INDIRIZZO DESTINAZIONE with "Accompag. con macchina attrezzata"
- **Row 9**: Compound names (Maria Grazia, Antonio Maria)
- **Row 10**: Hyphenated surname (Rossi-Bianchi)

**Expected Behavior**:
- 3 appointments should be filtered (ANNULLATO status)
- 4 appointments with "Accompag. con macchina attrezzata" should be highlighted (excluding cancelled ones)
- All Italian special characters should be preserved
- DESCRIZIONE PUNTO PARTENZA duplication should work correctly
- Conditional INDIRIZZO logic should work with empty INDIRIZZO DESTINAZIONE

**Test Coverage**:
- Punto Partenza duplication (Req 4.4)
- Cancelled appointment filtering (Req 4.3)
- Italian character preservation (Req 2.4)
- Yellow highlighting rule (Req 4.1)
- ASSISTITO column formation with compound names (Req 4.5)
- INDIRIZZO conditional concatenation (Req 4.6, 4.7)

---

### 3. malformed_input.csv

**Purpose**: Test error handling for missing required columns

**Contents**:
- CSV file missing the "DESCRIZIONE PUNTO PARTENZA" column
- Contains only 12 columns instead of the required 13

**Expected Behavior**:
- Should fail validation
- Should display Italian error message: "Il file CSV non contiene le colonne richieste: DESCRIZIONE PUNTO PARTENZA"
- Application should remain in usable state after error

**Test Coverage**:
- CSV validation (Req 2.3)
- Missing column error reporting (Req 9.3)
- Error recovery state (Req 9.6)
- Italian error messages (Req 8.3)

---

### 4. empty_input.csv

**Purpose**: Test handling of CSV file with no data rows

**Contents**:
- Header row only
- No data rows

**Expected Behavior**:
- Should parse successfully (valid structure)
- Should produce empty transformation result
- Should create new Excel sheet with only header and fissi data

**Test Coverage**:
- CSV parsing with empty data (Req 2.1)
- Transformation with no input rows (Req 4.1-4.9)
- Fissi data appending to empty sheet (Req 6.1, 6.2)

---

### 5. sample_workbook.xlsx

**Purpose**: Standard Excel workbook for testing sheet operations

**Structure**:
- **Sheet "1"**: Header = "05 gen 11 gen Settimana 1referente settimana = Marco Rossi 333-1234567"
- **Sheet "2"**: Header = "12 gen 18 gen Settimana 2referente settimana = Laura Bianchi 339-9876543"
- **Sheet "3"**: Header = "19 gen 25 gen Settimana 3referente settimana = Giuseppe Verdi 348-5551234"
- **Sheet "4"**: Header = "26 gen 01 feb Settimana 4referente settimana = Anna Ferrari 340-7778899"
  - Contains 3 sample data rows
- **Sheet "fissi"**: Contains 3 recurring appointments
  - Row 1: Header row (bold formatting)
  - Row 2: "Colombo Francesca" - Dialisi settimanale (yellow highlighting)
  - Row 3: "Greco Antonio" - Ritiro farmaci (normal formatting)
  - Row 4: "D'Angelo Nicolò" - Fisioterapia (italic formatting)

**Expected Behavior**:
- Should open successfully
- Should identify next sheet number as 5 (max numbered sheet + 1)
- Should locate "fissi" sheet successfully
- Should read header from sheet "4" correctly
- Should preserve all existing sheets when adding new sheet
- Should copy fissi data with formatting preserved

**Test Coverage**:
- Excel workbook opening (Req 3.1)
- Sheet name reading (Req 3.1)
- Sequential sheet numbering (Req 3.2)
- Sheet preservation (Req 3.3)
- Fissi sheet location (Req 3.4)
- Header parsing (Req 5.1, 5.2)
- Fissi data extraction (Req 6.1)
- Cell formatting preservation (Req 6.3)

---

## Usage in Tests

### Integration Tests

Use these files for end-to-end integration testing:

```csharp
[Test]
public void TestCompleteWorkflow_WithSampleData()
{
    // Arrange
    string csvPath = "TestData/sample_input.csv";
    string excelPath = "TestData/sample_workbook.xlsx";
    string outputPath = "TestData/output_test.xlsx";
    
    // Act
    var controller = new ApplicationController(...);
    controller.OnCSVFileSelected(csvPath);
    controller.OnExcelFileSelected(excelPath);
    controller.OnProcessButtonClicked();
    controller.OnDownloadButtonClicked(outputPath);
    
    // Assert
    // Verify output file exists and has correct structure
}
```

### Unit Tests

Use individual files for specific test scenarios:

```csharp
[Test]
public void TestCSVParser_WithEdgeCases()
{
    // Arrange
    var parser = new CSVParser();
    string csvPath = "TestData/edge_cases_input.csv";
    
    // Act
    var appointments = parser.ParseCSV(csvPath);
    
    // Assert
    Assert.AreEqual(10, appointments.Count);
    // Verify special characters preserved
    Assert.IsTrue(appointments.Any(a => a.NomeAssistito.Contains("Nicolò")));
}

[Test]
public void TestCSVValidation_WithMalformedFile()
{
    // Arrange
    var parser = new CSVParser();
    string csvPath = "TestData/malformed_input.csv";
    
    // Act & Assert
    Assert.Throws<InvalidOperationException>(() => parser.ParseCSV(csvPath));
}
```

### Property-Based Tests

Use these files as seed data for property-based test generators:

```csharp
// Feature: auser-excel-transformer, Property 2: Italian Character Preservation
[Property]
public Property ItalianCharacterPreservation()
{
    return Prop.ForAll<string>(text =>
    {
        // Use edge_cases_input.csv as reference for Italian characters
        var italianChars = "àèéìòùÀÈÉÌÒÙ'";
        // Test that parser preserves these characters
    });
}
```

---

## File Formats

### CSV Format

- **Delimiter**: Semicolon (;)
- **Encoding**: UTF-8 (to support Italian characters)
- **Line Endings**: Windows (CRLF)
- **Header**: Required, must match exact column names
- **Trailing Semicolon**: Present on each line

### Excel Format

- **Format**: .xlsx (Office Open XML)
- **Library**: EPPlus compatible
- **Sheets**: Numbered sheets (1, 2, 3, ...) and "fissi" sheet
- **Header**: First row of each numbered sheet
- **Data**: Starts from row 2

---

## Maintenance

When updating test data:

1. **Preserve Italian characters**: Ensure UTF-8 encoding is maintained
2. **Update README**: Document any new edge cases or scenarios
3. **Verify column structure**: CSV files must have all 13 required columns
4. **Test fissi sheet**: Ensure "fissi" sheet name is lowercase
5. **Validate dates**: Use Italian date format (DD/MM/YYYY)
6. **Check formatting**: Excel files should preserve cell formatting

---

## Expected Transformation Results

### From sample_input.csv

**Input**: 14 appointments
**After filtering ANNULLATO**: 11 appointments
**Yellow highlighted rows**: 4 rows (with "Accompag. con macchina attrezzata")
**Fissi rows appended**: 3 rows

**Total output rows**: 11 (transformed) + 3 (fissi) = 14 data rows (plus 1 header row)

### From edge_cases_input.csv

**Input**: 10 appointments
**After filtering ANNULLATO**: 7 appointments
**Yellow highlighted rows**: 4 rows (excluding cancelled ones)
**Fissi rows appended**: 3 rows

**Total output rows**: 7 (transformed) + 3 (fissi) = 10 data rows (plus 1 header row)

---

## Notes

- All test files use realistic Italian names, addresses, and locations
- Dates are in February 2026 to avoid conflicts with current dates
- Phone numbers in Excel headers are fictional
- Special characters test coverage includes: à è é ì ò ù ' -
- The "fissi" sheet demonstrates three types of formatting: yellow highlight, normal, and italic

