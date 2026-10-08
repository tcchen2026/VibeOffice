using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

// JSONL, one result for every argument. --baseline ORIGINAL SAVED compares diagnostic identities.
OpenXmlPackage Open(string path) => Path.GetExtension(path).ToLowerInvariant() switch {
    ".docx" or ".docm" or ".dotx" or ".dotm" => WordprocessingDocument.Open(path, false),
    ".xlsx" or ".xlsm" or ".xltx" or ".xltm" => SpreadsheetDocument.Open(path, false),
    ".pptx" or ".pptm" or ".ppsx" or ".ppsm" or ".potx" or ".potm" => PresentationDocument.Open(path, false),
    _ => throw new ArgumentException("Unsupported Office file variant: " + path),
};
List<Diagnostic> Validate(string path) {
    using var doc = Open(path);
    return new OpenXmlValidator(FileFormatVersions.Office2019) { MaxNumberOfErrors = 0 }.Validate(doc)
        .Select(e => new Diagnostic(e.Id, e.Part?.Uri.ToString() ?? "", e.Path?.XPath ?? "", e.ErrorType.ToString(), e.Description))
        .OrderBy(e => e.Part).ThenBy(e => e.Path).ThenBy(e => e.Id).ToList();
}
var json = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
if (args.FirstOrDefault() == "--stdin-pairs") {
    // Keep the SDK/JIT warm for large corpora. Each request gets exactly one independent result.
    string? line;
    while ((line = Console.ReadLine()) != null) {
        string file = "";
        try {
            using var request = JsonDocument.Parse(line);
            var original = request.RootElement.GetProperty("original").GetString()!;
            file = request.RootElement.GetProperty("saved").GetString()!;
            var prior = Validate(original).GroupBy(e => e.Key).ToDictionary(g => g.Key, g => g.Count());
            var diagnostics = Validate(file);
            var added = diagnostics.Where(e => {
                if (!prior.TryGetValue(e.Key, out var n) || n == 0) return true;
                prior[e.Key] = n - 1; return false;
            }).ToList();
            Console.WriteLine(JsonSerializer.Serialize(new { file, status = added.Count == 0 ? "ok" : "failed", diagnostics, newDiagnostics = added }, json));
        } catch (Exception e) {
            Console.WriteLine(JsonSerializer.Serialize(new { file, status = "failed", error = e.Message }, json));
        }
        Console.Out.Flush();
    }
    return 0;
}
var files = args;
Dictionary<string, int>? baseline = null;
if (args.FirstOrDefault() == "--baseline") {
    if (args.Length != 3) { Console.Error.WriteLine("Usage: oxval --baseline ORIGINAL SAVED"); return 2; }
    try { baseline = Validate(args[1]).GroupBy(e => e.Key).ToDictionary(g => g.Key, g => g.Count()); }
    catch (Exception e) { Console.WriteLine(JsonSerializer.Serialize(new { file = args[1], status = "failed", error = e.Message }, json)); return 1; }
    files = args.Skip(2).ToArray();
}
if (files.Length == 0) { Console.Error.WriteLine("Usage: oxval FILE... | --baseline ORIGINAL SAVED"); return 2; }
var failed = false;
foreach (var file in files) {
    try {
        var diagnostics = Validate(file);
        var added = diagnostics.Where(e => {
            if (baseline == null || !baseline.TryGetValue(e.Key, out var n) || n == 0) return true;
            baseline[e.Key] = n - 1; return false;
        }).ToList();
        failed |= added.Count != 0;
        Console.WriteLine(JsonSerializer.Serialize(new { file, status = added.Count == 0 ? "ok" : "failed", diagnostics, newDiagnostics = added }, json));
    } catch (Exception e) {
        failed = true;
        Console.WriteLine(JsonSerializer.Serialize(new { file, status = "failed", error = e.Message }, json));
    }
}
return failed ? 1 : 0;

record Diagnostic(string Id, string Part, string Path, string Type, string Description) {
    [System.Text.Json.Serialization.JsonIgnore] public string Key => JsonSerializer.Serialize(new[] { Part, Path, Id });
}
