public class DurationProcessor : IColumnProcessor
{
    private List<string> _durationBuffer = new List<string>();

    public void BufferText(string text)
    {
        _durationBuffer.Add(text);
    }

    public void Process(PdfData data)
    {
        if (_durationBuffer.Count > 0)
        {
            data.Duration = TimeSpan.Parse(string.Join("", _durationBuffer));
        }
    }

    public void Reset() => _durationBuffer.Clear();
}
