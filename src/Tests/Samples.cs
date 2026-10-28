public class Samples
{
    #region WordDocument

    [Test]
    public Task WordDocumentInstance()
    {
        using var document = WordDocument.Create();
        document.BuiltinDocumentProperties.Title = "The Title";
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

    #endregion

    #region PowerPointPresentation

    [Test]
    public Task PowerPointPresentationInstance()
    {
        using var presentation = PowerPointPresentation.Create();
        presentation.BuiltinDocumentProperties.Title = "The Title";
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
