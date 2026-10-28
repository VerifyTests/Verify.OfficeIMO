public class FormatTests
{
    #region VisioDocument

    [Test]
    public Task Visio()
    {
        var document = VisioDocument.Create();
        document.Title = "The Title";
        var page = document.AddPage("First");
        var shape = page.AddRectangle(1, 1, 2, 1);
        shape.Text = "Hello World";
        return Verify(document);
    }

    #endregion

    [Test]
    public Task Pdf()
    {
        var document = PdfDocument.Create(_ => _.Content(_ => _
                .H1("Hello World")
                .Paragraph(_ => _.Text("Second paragraph"))))
            .Meta(title: "The Title");
        return Verify(document);
    }

    [Test]
    public Task Odt()
    {
        var document = OdtDocument.Create();
        document.Metadata.Title = "The Title";
        document.AddParagraph("Hello World");
        document.AddParagraph("Second paragraph");
        return Verify(document);
    }

    [Test]
    public Task Ods()
    {
        var document = OdsDocument.Create();
        var sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("ID");
        sheet.Cell(0, 1).SetString("Name");
        sheet.Cell(1, 0).SetString("1");
        sheet.Cell(1, 1).SetString("John, Doe");
        return Verify(document);
    }

    [Test]
    public Task Rtf()
    {
        var document = RtfDocument.Create();
        document.Info.Title = "The Title";
        document.AddParagraph("Hello World");
        document.AddParagraph("Second paragraph");
        return Verify(document);
    }

    [Test]
    public Task Epub()
    {
        var publication = EpubPublication.Create("The Title", identifier: "urn:uuid:8b0c3c55-3a8f-4f4b-9d6c-1f0f6d6c1a11");
        publication.AddChapter("c1", "chapter1.xhtml", "Chapter One", "<p>Hello World</p>");
        publication.AddChapter("c2", "chapter2.xhtml", "Chapter Two", "<p>Second chapter</p>");
        return Verify(publication);
    }

    [Test]
    public Task Email()
    {
        var document = new EmailDocument
        {
            Subject = "The Subject",
            From = new("from@example.com", "From Name"),
            Date = new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.Zero),
            MessageId = "<message@example.com>"
        };
        document.Recipients.Add(new(EmailRecipientKind.To, new("to@example.com")));
        document.Body.Text = "Hello World";
        return Verify(document);
    }

    [Test]
    public Task Mhtml()
    {
        var document = new MhtmlDocument(
            "<html><body><h1>Hello World</h1><p>Second paragraph</p></body></html>",
            contentLocation: "https://example.com/page.html",
            subject: "The Subject");
        return Verify(document);
    }

    [Test]
    public Task OneNote()
    {
        var section = new OneNoteSection
        {
            Name = "Section"
        };
        var page = new OneNotePage
        {
            Title = "The Page"
        };
        var outline = new OneNoteOutline();
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new()
        {
            Text = "Hello World"
        });
        outline.Children.Add(paragraph);
        page.Outlines.Add(outline);
        section.Pages.Add(page);
        return Verify(section);
    }

    [Test]
    public Task Project()
    {
        using var document = ProjectDocument.Create();
        document.Title = "The Title";
        document.Settings.StartDate = new(2026, 10, 5, 8, 0, 0);
        document.Calendar = document.Calendars.AddStandardWorkingWeek();
        var summary = document.Tasks.AddSummary("Delivery");
        var design = summary.Children.Add("Design");
        design.Duration = ProjectDuration.WorkingDays(3);
        var build = summary.Children.Add("Build");
        build.Duration = ProjectDuration.WorkingDays(5);
        document.Dependencies.Add(design, build);
        return Verify(document);
    }

    [Test]
    public Task PublisherFile() =>
        VerifyFile("Samples/Simple.pub");

    [Test]
    public Task DjVuFile() =>
        VerifyFile("Samples/unicode.djvu");

    [Test]
    public Task AccessFile() =>
        VerifyFile("Samples/catalog-scaffold.accdb");

    [Test]
    public Task KeynoteFile() =>
        VerifyFile("Samples/keynote-created-v15.4.key");

    [Test]
    public Task XpsFile() =>
        VerifyFile("Samples/Test_Document.xps")
            .PagesToInclude(1);

    [Test]
    public Task ProjectFile() =>
        VerifyFile("Samples/delivery.mpp");
}
