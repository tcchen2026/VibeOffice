using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using System.Text;

// JSONL, one result for every argument. --baseline ORIGINAL SAVED compares diagnostic identities.
var version = FileFormatVersions.Office2019;
if (args.Length >= 2 && args[0] == "--version") {
    version = Enum.Parse<FileFormatVersions>(args[1], ignoreCase: true);
    args = args.Skip(2).ToArray();
}
OpenXmlPackage Open(string path) => Path.GetExtension(path).ToLowerInvariant() switch {
    ".docx" or ".docm" or ".dotx" or ".dotm" => WordprocessingDocument.Open(path, false),
    ".xlsx" or ".xlsm" or ".xltx" or ".xltm" => SpreadsheetDocument.Open(path, false),
    ".pptx" or ".pptm" or ".ppsx" or ".ppsm" or ".potx" or ".potm" => PresentationDocument.Open(path, false),
    _ => throw new ArgumentException("Unsupported Office file variant: " + path),
};
object Content(OpenXmlElement node) => new {
    name = node.NamespaceUri + ":" + node.LocalName,
    attributes = node.GetAttributes().Where(a => a.NamespaceUri != "http://www.w3.org/2000/xmlns/" && !(a.NamespaceUri == "http://schemas.openxmlformats.org/markup-compatibility/2006" && a.LocalName == "Ignorable"))
        .OrderBy(a => a.NamespaceUri).ThenBy(a => a.LocalName).Select(a => new[] { a.NamespaceUri, a.LocalName, a.Value }).ToArray(),
    text = node.ChildElements.Count == 0 ? node.InnerText : null,
    children = node.ChildElements.Select(Content).ToArray(),
};
Diagnostic Describe(ValidationErrorInfo error, Dictionary<OpenXmlElement, string> cache) {
    var path = error.Path?.XPath ?? "";
    var style = error.Node as Style ?? error.Node?.Ancestors<Style>().FirstOrDefault();
    // Styles keep their identity even when built-in definitions are inserted or
    // reordered. Compare the same source property, not its old list position.
    var stable = style?.StyleId?.Value is string id
        ? Regex.Replace(path, @"(/[^/:]+:style)\[\d+\]", m => m.Groups[1].Value + "[@styleId=" + JsonSerializer.Serialize(id) + "]")
        : null;
    // Regenerated runs may move an unchanged invalid source property. Match its
    // namespace-expanded XML as a second identity, retaining occurrence counts.
    string? fingerprint = null;
    if (error.Node != null && !cache.TryGetValue(error.Node, out fingerprint)) cache[error.Node] = fingerprint = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(Content(error.Node)))));
    return new Diagnostic(error.Id, error.Part?.Uri.ToString() ?? "", path, error.ErrorType.ToString(), error.Description, stable, fingerprint);
}
List<Diagnostic> Validate(string path) {
    using var doc = Open(path);
    var cache = new Dictionary<OpenXmlElement, string>();
    return new OpenXmlValidator(version) { MaxNumberOfErrors = 0 }.Validate(doc)
        .Select(e => Describe(e, cache))
        .OrderBy(e => e.Part).ThenBy(e => e.Path).ThenBy(e => e.Id).ToList();
}
var json = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
List<Diagnostic> Added(List<Diagnostic> prior, List<Diagnostic> saved) {
    var used = new bool[prior.Count];
    var matched = new bool[saved.Count];
    void Match(Func<Diagnostic, string?> key) {
        var buckets = prior.Select((d, i) => (key: key(d), i)).Where(x => x.key != null && !used[x.i])
            .GroupBy(x => x.key!).ToDictionary(g => g.Key, g => new Queue<int>(g.Select(x => x.i)));
        for (var i = 0; i < saved.Count; i++) {
            var k = key(saved[i]);
            if (matched[i] || k == null || !buckets.TryGetValue(k, out var queue) || queue.Count == 0) continue;
            used[queue.Dequeue()] = true; matched[i] = true;
        }
    }
    Match(d => d.Key); Match(d => d.ContentKey);
    return saved.Where((_, i) => !matched[i]).ToList();
}
if (args.FirstOrDefault() == "--comparison-self-test") {
    var a = new Diagnostic("schema", "/word/document.xml", "/p[1]/rPr[1]", "Schema", "invalid property", null, "same-content");
    var moved = a with { Path = "/p[2]/rPr[1]" };
    if (Added(new() { a }, new() { moved }).Count != 0 || Added(new() { a }, new() { a, moved }).Count != 1
        || Added(new() { a }, new() { moved with { ContentHash = "different" } }).Count != 1
        || Added(new() { a }, new() { a with { Description = "different error" } }).Count != 1)
        throw new Exception("Diagnostic comparison failed");
    Console.WriteLine("4 diagnostic identity/count checks passed"); return 0;
}
if (args.FirstOrDefault() == "--stdin-pairs") {
    // Keep the SDK/JIT warm for large corpora. Each request gets exactly one independent result.
    string? line;
    while ((line = Console.ReadLine()) != null) {
        string file = "";
        try {
            using var request = JsonDocument.Parse(line);
            var original = request.RootElement.GetProperty("original").GetString()!;
            file = request.RootElement.GetProperty("saved").GetString()!;
            var prior = Validate(original);
            var diagnostics = Validate(file);
            var added = Added(prior, diagnostics);
            Console.WriteLine(JsonSerializer.Serialize(new { file, status = added.Count == 0 ? "ok" : "failed", diagnostics, newDiagnostics = added }, json));
        } catch (Exception e) {
            Console.WriteLine(JsonSerializer.Serialize(new { file, status = "failed", error = e.Message }, json));
        }
        Console.Out.Flush();
    }
    return 0;
}
var files = args;
List<Diagnostic>? baseline = null;
if (args.FirstOrDefault() == "--baseline") {
    if (args.Length != 3) { Console.Error.WriteLine("Usage: oxval --baseline ORIGINAL SAVED"); return 2; }
    try { baseline = Validate(args[1]); }
    catch (Exception e) { Console.WriteLine(JsonSerializer.Serialize(new { file = args[1], status = "failed", error = e.Message }, json)); return 1; }
    files = args.Skip(2).ToArray();
}
if (files.Length == 0) { Console.Error.WriteLine("Usage: oxval FILE... | --baseline ORIGINAL SAVED"); return 2; }
var failed = false;
foreach (var file in files) {
    try {
        var diagnostics = Validate(file);
        var added = baseline == null ? diagnostics : Added(baseline, diagnostics);
        failed |= added.Count != 0;
        Console.WriteLine(JsonSerializer.Serialize(new { file, status = added.Count == 0 ? "ok" : "failed", diagnostics, newDiagnostics = added }, json));
    } catch (Exception e) {
        failed = true;
        Console.WriteLine(JsonSerializer.Serialize(new { file, status = "failed", error = e.Message }, json));
    }
}
return failed ? 1 : 0;

record Diagnostic(string Id, string Part, string Path, string Type, string Description, string? StablePath, string? ContentHash) {
    [System.Text.Json.Serialization.JsonIgnore] public string Key => JsonSerializer.Serialize(new[] { Part, StablePath ?? Path, Id, Description });
    [System.Text.Json.Serialization.JsonIgnore] public string? ContentKey => ContentHash == null ? null : JsonSerializer.Serialize(new[] { Part, Id, Description, ContentHash });
}
