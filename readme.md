# <img src="/src/icon.png" height="30px"> Verify.OfficeIMO

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://github.com/VerifyTests/Verify.OfficeIMO/actions/workflows/build.yml/badge.svg)](https://github.com/VerifyTests/Verify.OfficeIMO/actions/workflows/build.yml)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.OfficeIMO.svg)](https://www.nuget.org/packages/Verify.OfficeIMO/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of Word, Excel, and PowerPoint documents via [OfficeIMO](https://github.com/EvotecIT/OfficeIMO).<!-- singleLineInclude: intro. path: /docs/intro.include.md -->

**See [Milestones](../../milestones?state=closed) for release notes.**


## Sponsors


### Entity Framework Extensions<!-- include: sponsors. path: /docs/sponsors.include.md -->

[Entity Framework Extensions](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.OfficeIMO) is a major sponsor and is proud to contribute to the development this project.

[![Entity Framework Extensions](https://raw.githubusercontent.com/VerifyTests/Verify.OfficeIMO/refs/heads/main/docs/zzz.png)](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.OfficeIMO)

### Developed using JetBrains IDEs

[![JetBrains logo.](https://raw.githubusercontent.com/VerifyTests/Verify.OfficeIMO/main/docs/jetbrains.png)](https://jb.gg/OpenSourceSupport)<!-- endInclude -->


## NuGet

 * https://nuget.org/packages/Verify.OfficeIMO


## Usage


### Enable Verify.OfficeIMO

<!-- snippet: enable -->
<a id='snippet-enable'></a>
```cs
[ModuleInitializer]
public static void Initialize() =>
    VerifyOfficeIMO.Initialize();
```
<sup><a href='/src/Tests/ModuleInitializer.cs#L3-L9' title='Snippet source file'>snippet source</a> | <a href='#snippet-enable' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Outputs

| Type                     | Source | Info                                      | Derived               |
|--------------------------|--------|-------------------------------------------|-----------------------|
| `WordDocument`           | docx   | properties, sections, paragraphs, tables  | `txt` of all paragraphs |
| `ExcelDocument`          | xlsx   | properties, sheet names                   | `csv` per sheet       |
| `PowerPointPresentation` | pptx   | properties, slide count, hidden slides    | `txt` per slide (with notes) |

The source document is made deterministic via [DeterministicIoPackaging](https://github.com/SimonCropp/DeterministicIoPackaging). Streams and files with a `docx`, `xlsx`, or `pptx` extension are loaded via OfficeIMO and verified the same way.

What is left out is controlled by Verify's own settings: see [Leaving out what was derived](https://github.com/VerifyTests/Verify/blob/main/docs/paged-documents.md#leaving-out-what-was-derived) and [Excluding targets](https://github.com/VerifyTests/Verify/blob/main/docs/converter.md#excluding-targets). Sheets and slides are pages, so `PagesToInclude` limits them.


### WordDocument

<!-- snippet: WordDocument -->
<a id='snippet-WordDocument'></a>
```cs
[Test]
public Task WordDocument()
{
    using var document = OfficeIMO.Word.WordDocument.Create();
    document.BuiltinDocumentProperties.Title = "The Title";
    document.AddParagraph("Hello World");
    document.AddParagraph("Second paragraph");
    var table = document.AddTable(2, 2);
    table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
    table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";
    return Verify(document);
}
```
<sup><a href='/src/Tests/Samples.cs#L7-L22' title='Snippet source file'>snippet source</a> | <a href='#snippet-WordDocument' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### ExcelDocument

<!-- snippet: ExcelDocument -->
<a id='snippet-ExcelDocument'></a>
```cs
[Test]
public Task ExcelDocument()
{
    using var document = OfficeIMO.Excel.ExcelDocument.Create();
    document.BuiltinDocumentProperties.Title = "The Title";
    var sheet = document.AddWorksheet("Data");
    sheet.CellValue(1, 1, "ID");
    sheet.CellValue(1, 2, "Name");
    sheet.CellValue(2, 1, 1);
    sheet.CellValue(2, 2, "John, Doe");
    sheet.CellValue(3, 1, 2);
    sheet.CellValue(3, 2, "Jane Smith");
    var second = document.AddWorksheet("Second");
    second.CellValue(1, 1, "Other");
    return Verify(document);
}
```
<sup><a href='/src/Tests/Samples.cs#L24-L43' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcelDocument' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### PowerPointPresentation

<!-- snippet: PowerPointPresentation -->
<a id='snippet-PowerPointPresentation'></a>
```cs
[Test]
public Task PowerPointPresentation()
{
    using var presentation = OfficeIMO.PowerPoint.PowerPointPresentation.Create();
    presentation.BuiltinDocumentProperties.Title = "The Title";
    var slide = presentation.AddSlide();
    slide.AddTextBox("Hello World");
    slide.Notes.Text = "The notes";
    var second = presentation.AddSlide();
    second.AddTextBox("Second slide");
    return Verify(presentation);
}
```
<sup><a href='/src/Tests/Samples.cs#L45-L60' title='Snippet source file'>snippet source</a> | <a href='#snippet-PowerPointPresentation' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

