using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Text.Json;

const int defaultMaxErrors = 100;
const FileFormatVersions targetProfile = FileFormatVersions.Office2019;

if (args.Length > 0 && args[0] is "-h" or "--help")
{
    Console.WriteLine("Usage: dotnet run --project validation/openxml/Docxtor.OpenXmlValidation.csproj -- <file.docx> [--max-errors N]");
    return 0;
}
if (args.Length is < 1 or > 3 || args[0].StartsWith("-", StringComparison.Ordinal))
{
    Console.Error.WriteLine("Usage: dotnet run --project validation/openxml/Docxtor.OpenXmlValidation.csproj -- <file.docx> [--max-errors N]");
    return 2;
}

var maxErrors = defaultMaxErrors;
if (args.Length == 3 && args[1] == "--max-errors" &&
    int.TryParse(args[2], out var parsedMaxErrors) && parsedMaxErrors > 0)
{
    maxErrors = parsedMaxErrors;
}
else if (args.Length != 1)
{
    Console.Error.WriteLine("--max-errors must be a positive integer.");
    return 2;
}

var path = Path.GetFullPath(args[0]);
if (!File.Exists(path))
{
    Console.Error.WriteLine(JsonSerializer.Serialize(new { kind = "package_open_error", path, message = "File does not exist." }));
    return 2;
}

try
{
    using var document = WordprocessingDocument.Open(path, false);
    var validator = new OpenXmlValidator(targetProfile) { MaxNumberOfErrors = maxErrors };
    var findings = validator.Validate(document).Select(error => new
    {
        kind = "validation_finding",
        errorType = error.ErrorType.ToString(),
        errorId = error.Id,
        description = error.Description,
        partUri = error.Part?.Uri.ToString(),
        path = error.Path?.XPath,
    }).ToArray();

    foreach (var finding in findings)
        Console.WriteLine(JsonSerializer.Serialize(finding));

    Console.Error.WriteLine(JsonSerializer.Serialize(new
    {
        kind = "validation_summary",
        sdkVersion = typeof(OpenXmlValidator).Assembly.GetName().Version?.ToString(),
        targetProfile = targetProfile.ToString(),
        maxErrors,
        findingsReported = findings.Length,
        hasFindings = findings.Length > 0,
    }));
    return findings.Length == 0 ? 0 : 1;
}
catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or OpenXmlPackageException or ArgumentException)
{
    Console.Error.WriteLine(JsonSerializer.Serialize(new
    {
        kind = "package_open_error",
        path,
        exceptionType = exception.GetType().Name,
        message = exception.Message,
    }));
    return 2;
}
