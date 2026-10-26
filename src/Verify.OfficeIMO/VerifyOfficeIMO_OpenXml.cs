namespace VerifyTests;

public static partial class VerifyOfficeIMO
{
    public static bool OpenXmlRegistered { get; private set; }

    /// <summary>
    /// Verifies the Open XML SDK documents (WordprocessingDocument, SpreadsheetDocument, and
    /// PresentationDocument) as their OfficeIMO counterparts. Opt-in, since Verify.OpenXml registers
    /// converters for the same types: of two converters for one type, the one registered first is used.
    /// </summary>
    public static void RegisterOpenXml()
    {
        if (OpenXmlRegistered)
        {
            throw new("Already Registered");
        }

        OpenXmlRegistered = true;

        // OfficeIMO cannot take over an open package, so it loads a copy, which is disposed once
        // Verify is done with it. The document the test passed in is left as it was.
        VerifierSettings.RegisterFileConverter<WordprocessingDocument>(
            (document, context) => ConvertLoaded(WordDocument.Load(Copy(_ => document.Clone(_))), context));
        VerifierSettings.RegisterFileConverter<SpreadsheetDocument>(
            (document, context) => ConvertLoaded(ExcelDocument.Load(Copy(_ => document.Clone(_))), context));
        VerifierSettings.RegisterFileConverter<PresentationDocument>(
            (document, context) => ConvertLoaded(PowerPointPresentation.Load(Copy(_ => document.Clone(_))), context));
    }

    // Clone is only supported on the concrete document type, not on OpenXmlPackage
    static MemoryStream Copy(Func<Stream, IDisposable> clone)
    {
        var stream = new MemoryStream();
        clone(stream).Dispose();
        stream.Position = 0;
        return stream;
    }
}
