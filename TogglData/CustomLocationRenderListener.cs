using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

public class CustomLocationRenderListener : IEventListener
{
    private readonly List<IColumnProcessor> _columnProcessors;
    private readonly List<PdfData> _resultList;
    private float _lastX = 0;
    private bool _skipLines = false;

    public CustomLocationRenderListener(List<IColumnProcessor> columnProcessors, List<PdfData> resultList)
    {
        _columnProcessors = columnProcessors;
        _resultList = resultList;
    }

    public void EventOccurred(IEventData data, EventType type)
    {
        if (type == EventType.RENDER_TEXT)
        {
            var renderInfo = (TextRenderInfo)data;
            string text = renderInfo.GetText();
            var bottomLeft = renderInfo.GetDescentLine().GetStartPoint();
            float x = bottomLeft.Get(0);

            // Check for skipping lines
            if (_skipLines)
            {
                if (text == "TIME")
                {
                    _skipLines = false;
                    return;
                }
                return;
            }

            // Start skipping lines after finding "Detailed Report"
            if (text == "Detailed Report")
            {
                _skipLines = true;
                return;
            }

            if (text == "Created with toggl.com")
                return;

            if (x < _lastX)
            {
                ProcessCurrentLine();
            }

            BufferTextByPosition(x, text);
            _lastX = x;
        }
    }

    public ICollection<EventType> GetSupportedEvents()
    {
        return new HashSet<EventType> { EventType.RENDER_TEXT };
    }

    private void BufferTextByPosition(float x, string text)
    {
        if (x >= 30 && x < 221)
            _columnProcessors[0].BufferText(text); // Task
        else if (x >= 221 && x < 328)
            _columnProcessors[1].BufferText(text); // Tag
        else if (x >= 411 && x < 500)
            _columnProcessors[2].BufferText(text); // Duration
        else if (x >= 500 && x < 532)
            _columnProcessors[3].BufferText(text); // Time
    }

    private void ProcessCurrentLine()
    {
        var data = new PdfData();
        foreach (var processor in _columnProcessors)
        {
            processor.Process(data);
        }
        _resultList.Add(data);
        ResetProcessors();
    }

    private void ResetProcessors()
    {
        foreach (var processor in _columnProcessors)
        {
            processor.Reset();
        }
    }

    public void ProcessEndOfPage()
    {
        ProcessCurrentLine();
    }
}
