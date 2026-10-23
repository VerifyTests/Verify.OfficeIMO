// Creator, LastModifiedBy, Created and Modified are deliberately not surfaced: DeterministicIoPackaging
// strips them from docProps/core.xml (they are user/time-specific), so capturing them would make the
// info disagree with the deterministic document that Verify re-reads.
class DocumentInfo
{
    public string? Title { get; init; }
    public string? Subject { get; init; }
    public string? Description { get; init; }
    public string? Category { get; init; }
    public string? Keywords { get; init; }
    public string? Revision { get; init; }
    public string? Version { get; init; }
}

class WordInfo : DocumentInfo
{
    public required int Sections { get; init; }
    public required int Paragraphs { get; init; }
    public required int Tables { get; init; }
}

class ExcelInfo : DocumentInfo
{
    public required List<string> Sheets { get; init; }
}

class PowerPointInfo : DocumentInfo
{
    public required int SlideCount { get; init; }
    public List<int>? HiddenSlides { get; init; }
}
