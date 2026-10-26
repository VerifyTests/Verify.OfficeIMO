# <img src="/src/icon.png" height="30px"> Verify.OfficeIMO

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://github.com/VerifyTests/Verify.OfficeIMO/actions/workflows/build.yml/badge.svg)](https://github.com/VerifyTests/Verify.OfficeIMO/actions/workflows/build.yml)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.OfficeIMO.svg)](https://www.nuget.org/packages/Verify.OfficeIMO/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of Word, Excel, PowerPoint, Visio, PDF, XPS, OpenDocument, RTF, EPUB, Email, MHTML, OneNote, Project, Publisher, DjVu, iWork, and Access documents via [OfficeIMO](https://github.com/EvotecIT/OfficeIMO).<!-- singleLineInclude: intro. path: /docs/intro.include.md -->

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
<sup><a href='/src/Tests/ModuleInitializer.cs#L5-L11' title='Snippet source file'>snippet source</a> | <a href='#snippet-enable' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Outputs

| Type | Extensions | Source | Derived |
|------|------------|--------|---------|
| `WordDocument` | docx | docx | `txt` of all paragraphs, `png` per page |
| `ExcelDocument` | xlsx | xlsx | `csv` and `png` per sheet |
| `PowerPointPresentation` | pptx | pptx | `txt` (with notes) and `png` per slide |
| `VisioDocument` | vsdx | vsdx | `txt` and `png` per page |
| `PdfDocument`, `PdfReadDocument` | pdf | pdf (authored only) | `txt` and `png` per page |
| `XpsDocument` | xps, oxps | xps | `txt` and `png` per page |
| `OdfDocument` | odt, ods, odp, odg | same as loaded | odt: `txt`; ods: `csv` per sheet; odp: `txt` per slide |
| `RtfDocument` | rtf | rtf | `txt` of all paragraphs |
| `EpubPublication`, `EpubDocument` | epub | epub (authored only) | `txt` and `png` per chapter |
| `EmailDocument` | eml, msg | eml | `txt`, `html`, and `png` of the body |
| `MhtmlDocument` | mht, mhtml | mht | `html` and `png` of the page |
| `OneNoteSection` | one | none | `txt` and `png` per page |
| `ProjectDocument` | mpp | Project xml | `txt` of the task tree, `png` per page of the Gantt view |
| `PublisherDocument` | pub | none | `txt` and `png` per page |
| `DjVuDocument` | djvu | none | `txt` and `png` per page |
| `IWorkSourceDocument` | pages, numbers, key | none | Pages: `txt` and `csv` per table; Numbers: `csv` per table; Keynote: `txt` per slide |
| `AccessDocument` | accdb, mdb | none | `csv` per table |

Each also has an info file of its properties and structure.

Source is "none" where OfficeIMO reads the format but does not write it (Publisher, DjVu, iWork), or does not write it deterministically (OneNote, Access). Office packages (docx, xlsx, pptx, vsdx, xps, epub) are made deterministic via [DeterministicIoPackaging](https://github.com/SimonCropp/DeterministicIoPackaging). An `EpubPublication` is saved with `dcterms:modified` pinned to 2000-01-01. A Project schedule is only rendered when the project has a calendar.

Images are rendered by OfficeIMO itself, with no Office install needed. OpenDocument and RTF have no `png`, since OfficeIMO renders them only via its pdf packages. `ExcludeDerivedTargets("png")` skips rendering.

Streams and files with any of the extensions above are loaded via OfficeIMO and verified the same way. Other Verify plugins that register the same extensions (for example `pdf` or `xlsx`) will conflict.

What is left out is controlled by Verify's own settings: see [Leaving out what was derived](https://github.com/VerifyTests/Verify/blob/main/docs/paged-documents.md#leaving-out-what-was-derived) and [Excluding targets](https://github.com/VerifyTests/Verify/blob/main/docs/converter.md#excluding-targets). Pages, sheets, slides, and chapters are pages, so `PagesToInclude` limits them.


### Fonts

A png is rendered with the fonts installed on the machine. Office's default font, Aptos, is installed only with Office, so a build server without Office renders a fallback, and the png differs from one made on a developer machine. Fonts added to `VerifyOfficeIMO.Fonts` are used ahead of installed ones, so registering the fonts the documents use makes the png the same everywhere:

<!-- snippet: Fonts -->
<a id='snippet-Fonts'></a>
```cs
// Aptos is Office's default font, but is installed only with Office, so a build server renders a
// fallback. Registering it makes the png the same everywhere.
[ModuleInitializer]
public static void InitializeFonts()
{
    var directory = Path.Combine(ProjectFiles.ProjectDirectory, "..", "Fonts");
    AddFont(directory, "Aptos", "Aptos_400.ttf", OfficeFontStyle.Regular);
    AddFont(directory, "Aptos", "Aptos_400_Italic.ttf", OfficeFontStyle.Italic);
    AddFont(directory, "Aptos", "Aptos_700.ttf", OfficeFontStyle.Bold);
    AddFont(
        directory,
        "Aptos",
        "Aptos_700_Italic.ttf",
        OfficeFontStyle.Bold | OfficeFontStyle.Italic);
    AddFont(directory, "Aptos Display", "Aptos_Display_700.ttf", OfficeFontStyle.Bold);
}

static void AddFont(string directory, string family, string file, OfficeFontStyle style) =>
    VerifyOfficeIMO.Fonts.Add(family, File.ReadAllBytes(Path.Combine(directory, file)), style);
```
<sup><a href='/src/Tests/ModuleInitializer.cs#L21-L43' title='Snippet source file'>snippet source</a> | <a href='#snippet-Fonts' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Open XML SDK documents

`WordprocessingDocument`, `SpreadsheetDocument`, and `PresentationDocument` from the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) can be verified as their OfficeIMO counterparts, giving the same outputs as `WordDocument`, `ExcelDocument`, and `PowerPointPresentation`. OfficeIMO cannot take over an open package, so the document is copied and the copy loaded, leaving the one passed in as it was. This is opt-in, since [Verify.OpenXml](https://github.com/VerifyTests/Verify.OpenXml) converts the same types:

<!-- snippet: RegisterOpenXml -->
<a id='snippet-RegisterOpenXml'></a>
```cs
[ModuleInitializer]
public static void InitializeOpenXml() =>
    VerifyOfficeIMO.RegisterOpenXml();
```
<sup><a href='/src/Tests/ModuleInitializer.cs#L13-L19' title='Snippet source file'>snippet source</a> | <a href='#snippet-RegisterOpenXml' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

<!-- snippet: WordprocessingDocument -->
<a id='snippet-WordprocessingDocument'></a>
```cs
[Test]
public Task WordprocessingDocument()
{
    using var document = OpenWord();
    return Verify(document);
}
```
<sup><a href='/src/Tests/OpenXmlTests.cs#L7-L16' title='Snippet source file'>snippet source</a> | <a href='#snippet-WordprocessingDocument' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Using with Verify.OpenXml

Both packages convert Word, Excel, and PowerPoint, so when both are initialized, which one converts a document depends on the order they are initialized in. Verify resolves the two kinds of converter differently:

 * **Types** (`WordprocessingDocument`, `SpreadsheetDocument`, `PresentationDocument`): the converter registered **first** is used.
 * **Extensions** (`docx`, `xlsx`, `pptx`, as streams or with `VerifyFile`): the converter registered **last** is used.

`VerifyOpenXml.Initialize` registers both kinds, `VerifyOfficeIMO.Initialize` registers the extensions, and `VerifyOfficeIMO.RegisterOpenXml` registers the types. Separate `[ModuleInitializer]` methods give no order that is clear from the code, so call them from one method, in the order that gives the split wanted.

Verify.OpenXml for the Open XML SDK types and the extensions; Verify.OfficeIMO for the OfficeIMO types and every other format:

```cs
[ModuleInitializer]
public static void Initialize()
{
    VerifyOfficeIMO.Initialize();
    VerifyOpenXml.Initialize();
}
```

Verify.OfficeIMO for everything; Verify.OpenXml registered but never used for these three formats:

```cs
[ModuleInitializer]
public static void Initialize()
{
    VerifyOfficeIMO.RegisterOpenXml();
    VerifyOpenXml.Initialize();
    VerifyOfficeIMO.Initialize();
}
```


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
<sup><a href='/src/Tests/Samples.cs#L3-L18' title='Snippet source file'>snippet source</a> | <a href='#snippet-WordDocument' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L20-L39' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcelDocument' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L41-L56' title='Snippet source file'>snippet source</a> | <a href='#snippet-PowerPointPresentation' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

