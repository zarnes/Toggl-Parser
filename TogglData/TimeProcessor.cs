
using System.Globalization;

public class TimeProcessor : IColumnProcessor
{
    private List<string> _timeBuffer = new List<string>();

    public void BufferText(string text)
    {
        _timeBuffer.Add(text);
    }

    public void Process(PdfData data)
    {
        if (_timeBuffer.Count >= 3) // Attendre 3 éléments pour traiter le temps
        {
            string timeStr = _timeBuffer[0] + " " + _timeBuffer[1];
            string dateStr = _timeBuffer[2];

            var times = timeStr.Split('-');
            if (times.Length == 2)
            {
                DateTime date = DateTime.ParseExact(dateStr, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                DateTime startTime = DateTime.ParseExact(times[0].Trim() + " " + date.ToString("MM/dd/yyyy"), "hh:mm tt MM/dd/yyyy", CultureInfo.InvariantCulture);
                DateTime endTime = DateTime.ParseExact(times[1].Trim() + " " + date.ToString("MM/dd/yyyy"), "hh:mm tt MM/dd/yyyy", CultureInfo.InvariantCulture);

                data.StartTime = startTime;
                data.EndTime = endTime;
            }
        }
    }

    public void Reset() => _timeBuffer.Clear();
}
