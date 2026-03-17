using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var filePath = args[0];

// Check raw XML inside the xlsx zip for sheet 6 formulas
Console.WriteLine("=== RAW XML CHECK (sheet '6') ===");
using (var zip = ZipFile.OpenRead(filePath))
{
    // Find sheet6 - list all sheet entries
    var sheetEntries = zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/")).ToList();
    foreach (var e in sheetEntries)
        Console.WriteLine($"  Entry: {e.FullName}");

    // Also check workbook.xml to map sheet names to rIds
    var wbEntry = zip.GetEntry("xl/workbook.xml");
    if (wbEntry != null)
    {
        using var stream = wbEntry.Open();
        var doc = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Console.WriteLine("\n=== WORKBOOK SHEETS ===");
        foreach (var sh in doc.Descendants(ns + "sheet"))
            Console.WriteLine($"  name='{sh.Attribute("name")?.Value}' r:id='{sh.Attribute("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")?.Value}'");
    }

    // Find the sheet with name "6" via relationships
    var relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
    string? sheet6Target = null;
    if (relsEntry != null)
    {
        using var stream = relsEntry.Open();
        var doc = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        // Get workbook sheet rIds
        var wbEntry2 = zip.GetEntry("xl/workbook.xml");
        if (wbEntry2 != null)
        {
            using var wbStream = wbEntry2.Open();
            var wbDoc = XDocument.Load(wbStream);
            XNamespace wbNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var sheet6El = wbDoc.Descendants(wbNs + "sheet").FirstOrDefault(s => s.Attribute("name")?.Value == "6");
            var rId = sheet6El?.Attribute("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")?.Value;
            if (rId != null)
            {
                var rel = doc.Descendants(ns + "Relationship").FirstOrDefault(r => r.Attribute("Id")?.Value == rId);
                sheet6Target = rel?.Attribute("Target")?.Value;
                Console.WriteLine($"\nSheet '6' target: {sheet6Target}");
            }
        }
    }

    if (sheet6Target != null)
    {
        var entry = zip.GetEntry($"xl/{sheet6Target}") ?? zip.GetEntry(sheet6Target);
        if (entry != null)
        {
            using var stream = entry.Open();
            var doc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            Console.WriteLine("\n=== ROWS 3-8, COL 4 AND 6 FORMULAS ===");
            foreach (var row in doc.Descendants(ns + "row").Where(r => {
                var attr = r.Attribute("r");
                if (attr == null) return false;
                return int.TryParse(attr.Value, out int n) && n >= 3 && n <= 8;
            }))
            {
                var rowNum = row.Attribute("r")?.Value;
                // Find cells in col D (col 4) and F (col 6)
                var c4 = row.Descendants(ns + "c").FirstOrDefault(c => c.Attribute("r")?.Value?.StartsWith("D") == true);
                var c6 = row.Descendants(ns + "c").FirstOrDefault(c => c.Attribute("r")?.Value?.StartsWith("F") == true);
                var f4 = c4?.Element(ns + "f")?.Value ?? "(no formula)";
                var v4 = c4?.Element(ns + "v")?.Value ?? "(no value)";
                var f6 = c6?.Element(ns + "f")?.Value ?? "(no formula)";
                var v6 = c6?.Element(ns + "v")?.Value ?? "(no value)";
                Console.WriteLine($"  Row {rowNum}: D formula='{f4}' val='{v4}' | F formula='{f6}' val='{v6}'");
            }
        }
    }
}
