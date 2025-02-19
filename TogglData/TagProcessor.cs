public class TagProcessor : IColumnProcessor
{
    private List<string> _tagBuffer = new List<string>();

    public void BufferText(string text)
    {
        _tagBuffer.Add(text);
    }

    public void Process(PdfData data)
    {
        if (_tagBuffer.Count > 0)
        {
            data.Tag = string.Join(" ", _tagBuffer);
        }
    }

    public void Reset() => _tagBuffer.Clear();
}
