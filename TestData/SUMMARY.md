# Test Data Files Summary

## Created Files

### CSV Files (4 files)

1. **sample_input.csv** - Standard test case
   - 14 appointments with realistic Italian data
   - Mix of activities and statuses
   - Tests normal workflow

2. **edge_cases_input.csv** - Edge cases and special scenarios
   - 10 appointments with edge cases
   - Special characters (àèéìòù, apostrophes)
   - Compound names and hyphenated surnames
   - Empty fields to test conditional logic

3. **malformed_input.csv** - Error handling test
   - Missing required column (DESCRIZIONE PUNTO PARTENZA)
   - Tests validation and error messages

4. **empty_input.csv** - Empty data test
   - Header only, no data rows
   - Tests handling of empty input

### Excel File (1 file)

5. **sample_workbook.xlsx** - Standard Excel workbook
   - 4 numbered sheets (1, 2, 3, 4)
   - Sheet 4 header: "26 gen 01 feb Settimana 4referente settimana = Anna Ferrari 340-7778899"
   - "fissi" sheet with 3 recurring appointments
   - Demonstrates formatting preservation (bold, yellow, italic)

### Documentation (2 files)

6. **README.md** - Comprehensive documentation
   - Detailed description of each test file
   - Expected behaviors and test coverage
   - Usage examples for unit and integration tests
   - Maintenance guidelines

7. **SUMMARY.md** - This file
   - Quick overview of all test data files

## Test Coverage

These files validate:

- **Requirement 2.1**: CSV data parsing with all required columns ✓
- **Requirement 2.2**: Column extraction ✓
- **Requirement 2.3**: Malformed data error handling ✓
- **Requirement 2.4**: Italian character preservation ✓
- **Requirement 3.1**: Excel sheet name reading ✓
- **Requirement 3.2**: Sequential sheet numbering ✓
- **Requirement 3.3**: Sheet preservation ✓
- **Requirement 3.4**: Fissi sheet location ✓
- **Requirement 4.1**: Yellow highlighting rule ✓
- **Requirement 4.3**: Cancelled appointment filtering ✓
- **Requirement 4.4**: Punto Partenza duplication ✓
- **Requirement 4.5**: ASSISTITO column formation ✓
- **Requirement 4.6-4.7**: INDIRIZZO conditional concatenation ✓
- **Requirement 5.1-5.2**: Header parsing ✓
- **Requirement 6.1**: Fissi data extraction ✓
- **Requirement 6.3**: Cell formatting preservation ✓
- **Requirement 8.3**: Italian error messages ✓
- **Requirement 9.3**: Missing column error reporting ✓
- **Requirement 9.6**: Error recovery state ✓

## Quick Start

### Run Integration Test
```csharp
string csvPath = "TestData/sample_input.csv";
string excelPath = "TestData/sample_workbook.xlsx";
// Process and verify output
```

### Test Edge Cases
```csharp
string csvPath = "TestData/edge_cases_input.csv";
// Verify special character handling, filtering, etc.
```

### Test Error Handling
```csharp
string csvPath = "TestData/malformed_input.csv";
// Should throw exception with Italian error message
```

## File Characteristics

- **Encoding**: UTF-8 (supports Italian characters)
- **CSV Delimiter**: Semicolon (;)
- **Date Format**: DD/MM/YYYY
- **Excel Format**: .xlsx (EPPlus compatible)
- **Italian Characters**: àèéìòù, apostrophes, hyphens
- **Realistic Data**: Italian names, addresses, cities

## Expected Results

### sample_input.csv
- Input: 14 appointments
- After filtering: 11 appointments (3 ANNULLATO removed)
- Yellow highlighted: 4 rows
- With fissi: 14 total data rows

### edge_cases_input.csv
- Input: 10 appointments
- After filtering: 7 appointments (3 ANNULLATO removed)
- Yellow highlighted: 4 rows
- With fissi: 10 total data rows

### malformed_input.csv
- Should fail validation
- Error message: "Il file CSV non contiene le colonne richieste: DESCRIZIONE PUNTO PARTENZA"

### empty_input.csv
- Input: 0 appointments
- After filtering: 0 appointments
- With fissi: 3 total data rows (only fissi)

