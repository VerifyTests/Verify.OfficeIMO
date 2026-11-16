public class Samples
{
    #region WordDocument

    [Test]
    public Task WordDocumentInstance()
    {
        using var document = WordDocument.Create();
        document.Properties.Title = "The Title";
        document.AddParagraph("Hello World");
        document.AddParagraph("Second paragraph");
        var table = document.AddTable(2, 2);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
        table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";
        return Verify(document);
    }

    #endregion

    #region ExcelDocument

    [Test]
    public Task ExcelDocumentInstance()
    {
        using var document = ExcelDocument.Create();
        document.Properties.Title = "The Title";
        var sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("ID");
        sheet.Cell(1, 2).SetValue("Name");
        sheet.Cell(2, 1).SetValue(1);
        sheet.Cell(2, 2).SetValue("John, Doe");
        sheet.Cell(3, 1).SetValue(2);
        sheet.Cell(3, 2).SetValue("Jane Smith");
        var second = document.AddSheet("Second");
        second.Cell(1, 1).SetValue("Other");
        return Verify(document);
    }

    #endregion

    #region PowerPointDocument

    [Test]
    public Task PowerPointDocumentInstance()
    {
        using var presentation = PowerPointDocument.Create();
        presentation.Properties.Title = "The Title";
        var slide = presentation.AddSlide();
        slide.AddTextBox("Hello World");
        slide.Notes.Text = "The notes";
        var second = presentation.AddSlide();
        second.AddTextBox("Second slide");
        return Verify(presentation);
    }

    #endregion

    [Test]
    public Task ExcludeDerived() =>
        Verify(BuildWord())
            .ExcludeDerivedTargets("txt");

    [Test]
    public async Task VerifyDocxFile()
    {
        using var stream = new MemoryStream();
        await using (var document = BuildWord())
        {
            await document.SaveAsync(stream);
        }

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    static WordDocument BuildWord()
    {
        var document = WordDocument.Create();
        document.AddParagraph("Hello World");
        return document;
    }
}
