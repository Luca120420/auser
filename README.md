# Auser Excel Transformer

A Windows desktop application that transforms weekly CSV data exports into formatted Excel sheets for Auser (Associazione per l'autogestione dei servizi e la solidarietà), an Italian non-profit supporting elderly care services.

## Project Structure

```
AuserExcelTransformer/
├── Models/              # Data model classes
├── Services/            # Business logic and service classes
├── UI/                  # User interface components
├── Tests/               # Unit tests and property-based tests
├── Properties/          # Application properties and resources
│   ├── Resources.it.resx           # Italian resource strings
│   └── Resources.it.Designer.cs    # Auto-generated resource accessor
├── Program.cs           # Application entry point
└── AuserExcelTransformer.csproj    # Project file
```

## Dependencies

The project uses the following NuGet packages:

- **EPPlus 7.0.5** - Excel file manipulation (.xlsx)
- **CsvHelper 30.0.1** - CSV file parsing
- **FsCheck 2.16.6** - Property-based testing framework
- **NUnit 4.0.1** - Unit testing framework
- **NUnit3TestAdapter 4.5.0** - NUnit test adapter for Visual Studio
- **Microsoft.NET.Test.Sdk 17.8.0** - .NET testing SDK

## Requirements

- .NET 6.0 or later
- Windows 10 or Windows 11
- No Microsoft Excel installation required

## Building the Project

To build the project, run:

```bash
dotnet build
```

To run tests:

```bash
dotnet test
```

To run the application:

```bash
dotnet run
```

## Localization

All user-facing text is in Italian. Strings are stored in `Properties/Resources.it.resx` for easy maintenance and future localization.

## License

This application is developed for Auser, an Italian non-profit organization.
