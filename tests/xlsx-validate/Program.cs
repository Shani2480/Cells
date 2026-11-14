// Validates .xlsx files with the Open XML SDK (schema + semantic checks).
// Usage: dotnet run -- [--baseline original.xlsx] file1.xlsx [file2.xlsx ...]   Exit code 1 if any file has errors.
// --baseline: errors the original file has too (Excel's own parts, kept byte for byte) are not counted.
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var files = new List<string>();
string? baseline = null;
for (var i = 0; i < args.Length; i++)
{
    if (args[i] == "--baseline" && i + 1 < args.Length) baseline = args[++i];
    else files.Add(args[i]);
}
if (files.Count == 0)
{
    Console.Error.WriteLine("usage: xlsx-validate [--baseline original.xlsx] <file.xlsx> [...]");
    return 2;
}

var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
string Key(ValidationErrorInfo e) => $"{e.Part?.Uri}|{e.Path?.XPath}|{e.Description}";
// Errors as text, read while the package is still open.
List<(string Key, string Where, string Description, string Type)> Validate(string path)
{
    using var doc = SpreadsheetDocument.Open(path, false);
    return validator.Validate(doc)
        .Select(e => (Key(e), $"{e.Part?.Uri} {e.Path?.XPath}", e.Description, e.ErrorType.ToString()))
        .ToList();
}

var known = new HashSet<string>();
if (baseline != null) foreach (var e in Validate(baseline)) known.Add(e.Key);

var failed = false;
foreach (var path in files)
{
    try
    {
        var all = Validate(path);
        var errors = all.Where(e => !known.Contains(e.Key)).ToList();
        var skipped = all.Count - errors.Count;
        Console.WriteLine($"{Path.GetFileName(path)}: {errors.Count} error(s)" + (skipped > 0 ? $" ({skipped} also in the baseline)" : ""));
        foreach (var e in errors.Take(50))
        {
            Console.WriteLine($"  [{e.Type}] {e.Where}");
            Console.WriteLine($"      {e.Description}");
        }
        failed |= errors.Count > 0;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"{Path.GetFileName(path)}: cannot open package: {ex.Message}");
        failed = true;
    }
}
return failed ? 1 : 0;
