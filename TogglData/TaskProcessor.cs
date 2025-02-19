public class TaskProcessor : IColumnProcessor
{
    private List<string> _taskBuffer = new List<string>();

    public void BufferText(string text)
    {
        _taskBuffer.Add(text);
    }

    public void Process(PdfData data)
    {
        if (_taskBuffer.Count > 0)
        {
            if (_taskBuffer.Count >= 3)
            {
                data.Client = _taskBuffer[^1];
                data.Project = _taskBuffer[^2];
            }
            else
            {
                data.Project = _taskBuffer[^1];
            }

            data.Description = string.Join(" ", _taskBuffer.Take(_taskBuffer.Count - 1));
        }
    }

    public void Reset() => _taskBuffer.Clear();
}
