public class OpenXmlTests
{
    #region WordprocessingDocument

    [Test]
    public Task WordprocessingDocument()
    {
        using var document = OpenWord();
        return Verify(document);
    }

    #endregion

    [Test]
    public Task SpreadsheetDocument()
    {
        using var stream = new MemoryStream();
        using (var source = ExcelDocument.Create())
        {
            var sheet = source.AddWorksheet("Data");
            sheet.CellValue(1, 1, "ID");
            sheet.CellValue(1, 2, "Name");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(2, 2, "John Doe");
            source.Save(stream);
        }

        stream.Position = 0;
        using var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(stream, false);
        return Verify(document);
    }

    [Test]
    public Task PresentationDocument()
    {
        using var stream = new MemoryStream();
        using (var source = PowerPointPresentation.Create())
        {
            source.AddSlide().AddTextBox("Hello World");
            source.Save(stream);
        }

        stream.Position = 0;
        using var document = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(stream, false);
        return Verify(document);
    }

    // the document passed in is only copied, so is still usable after verifying
    [Test]
    public async Task DocumentStillUsable()
    {
        using var document = OpenWord();
        await Verify(document);
        await Assert.That(document.MainDocumentPart!.Document!.Body!.InnerText).Contains("Hello World");
    }

    static WordprocessingDocument OpenWord()
    {
        var stream = new MemoryStream();
        using (var source = WordDocument.Create())
        {
            source.AddParagraph("Hello World");
            source.Save(stream);
        }

        stream.Position = 0;
        return DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
    }
}
