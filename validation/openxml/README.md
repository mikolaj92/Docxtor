# Optional Open XML SDK validator

This independent validator runs outside Docxtor's Python runtime and never
changes the document. It pins .NET SDK 8.0.408 and
`DocumentFormat.OpenXml` 3.3.0; the validation target is
`FileFormatVersions.Office2019`.

```sh
dotnet run --project validation/openxml/Docxtor.OpenXmlValidation.csproj -- file.docx
# Optional positive cap; default: 100 findings
dotnet run --project validation/openxml/Docxtor.OpenXmlValidation.csproj -- file.docx --max-errors 25
```

Each finding is one JSON line on stdout, with a final summary on stderr. Exit
codes: `0` means no findings, `1` means at least one conformance finding, and
`2` means bad invocation or package-open/validator execution failure. A missing
SDK is a command failure, never a pass. The error cap bounds findings reported
by the SDK; it does not mean the package has no additional findings.

This validates OOXML conformance against the selected SDK profile. It does not
assert visual fidelity, business semantics, or every behavior of Microsoft
Office. It complements, not replaces, Docxtor's own package checks.
