
using System.Globalization;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using SpreadsheetLight;

public class PdfData
{
    public string Description { get; set; }
    public string Client { get; set; }
    public string Project { get; set; }
    public string Tag { get; set; }
    public TimeSpan Duration { get; set; }
    public DateTime StartTime { get; set; }
    public DateTime EndTime { get; set; }
}

public interface IColumnProcessor
{
    void BufferText(string text);
    void Process(PdfData data);
    void Reset();
}

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

public class PdfDataExtractor
{
    private readonly List<IColumnProcessor> _columnProcessors;

    public PdfDataExtractor(List<IColumnProcessor> columnProcessors)
    {
        _columnProcessors = columnProcessors;
    }

    public List<PdfData> Extract(string pdfPath)
    {
        var resultList = new List<PdfData>();

        using (PdfReader reader = new PdfReader(pdfPath))
        using (PdfDocument pdfDoc = new PdfDocument(reader))
        {
            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
            {
                var listener = new CustomLocationRenderListener(_columnProcessors, resultList);
                PdfCanvasProcessor processor = new PdfCanvasProcessor(listener);
                processor.ProcessPageContent(pdfDoc.GetPage(page));

                // Traiter les données à la fin de la page
                listener.ProcessEndOfPage();
            }
        }

        return resultList;
    }
}

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

public class ExcelGenerator
{
    public List<PdfData> Data { get; }
    public string Filepath { get; }
    public double TimeBase { get; }
    public double TimeFacturable { get; }

    public ExcelGenerator(List<PdfData> data, double timeBase, double timeFacturable, string filepath)
    {
        TimeBase = timeBase;
        TimeFacturable = timeFacturable;
        Data = data;
        Filepath = filepath;
    }

    public void CreateExcel()
    {
        SLDocument doc = new SLDocument();
        int column = 1;

        SLStyle headerStyle = new SLStyle();
        headerStyle.SetFontBold(true);
        doc.SetRowStyle(1, headerStyle);

        doc.SetColumnWidth(column++, 35);
        doc.SetColumnWidth(column++, 60);
        doc.SetColumnWidth(column++, 20);
        doc.SetColumnWidth(column++, 20);
        doc.SetColumnWidth(column++, 10);
        doc.SetColumnWidth(column++, 10);
        doc.SetColumnWidth(column++, 10);
        doc.SetColumnWidth(column++, 10);
        doc.SetColumnWidth(column++, 20);
        doc.SetColumnWidth(column++, 40);

        column = 1;

        doc.SetCellValue(1, column++, "Projet");
        doc.SetCellValue(1, column++, "Description");
        doc.SetCellValue(1, column++, "Début");
        doc.SetCellValue(1, column++, "Fin");
        doc.SetCellValue(1, column++, "Semaine");
        doc.SetCellValue(1, column++, "RMA");
        doc.SetCellValue(1, column++, "Mois");
        doc.SetCellValue(1, column++, "Durée (h), base de " + TimeBase + "h");
        doc.SetCellValue(1, column++, "Durée pondérée sur " + TimeFacturable + "h");
        doc.SetCellValue(1, column++, "Tags");

        double timeMultiplier = TimeFacturable / TimeBase;
        Calendar calendar = CultureInfo.InvariantCulture.Calendar;
        int[] rmas = [1, 5, 9, 14, 18, 22, 26, 29, 35, 39, 44, 48, 54];

        for (int index = 0; index <  Data.Count; ++index)
        {
            int row = index + 2;
            column = 1;
            PdfData line = Data[index];
            doc.SetCellValue(row, column++, line.Project);
            doc.SetCellValue(row, column++, line.Description);

            doc.SetCellValue(row, column++, line.StartTime.ToString());
            doc.SetCellValue(row, column++, line.EndTime.ToString());

            int weekOfYear = calendar.GetWeekOfYear(line.StartTime, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            doc.SetCellValueNumeric(row, column++, weekOfYear.ToString());
            int rma = 12;
            for (int i = 0; i < rmas.Length; ++i)
            {
                if (rmas[i] > weekOfYear)
                {
                    rma = i;
                    break;
                }
            }
            doc.SetCellValueNumeric(row, column++, rma.ToString());

            doc.SetCellValueNumeric(row, column++, line.StartTime.Month.ToString());

            double hours = Math.Round(line.Duration.TotalMinutes / 60, 2);
            string hoursString = hours.ToString().Replace(',', '.');
            doc.SetCellValueNumeric(row, column++, hoursString);

            double hoursMultiplied = Math.Round((line.Duration.TotalMinutes * timeMultiplier) / 60, 2);
            string hoursMultipliedString = hoursMultiplied.ToString().Replace(',', '.');
            doc.SetCellValueNumeric(row, column++, hoursMultipliedString);

            string[] tags = line.Tag.Split(',');
            foreach(string tag in tags.Reverse())
            {
                doc.SetCellValue(row, column++, tag);
            }
        }

        doc.SaveAs(Filepath);
    }
}

public class Program
{
    static float HOURS_PER_DAY_WORKED = 7;
    static float HOURS_PER_DAY_BILLABLE = 8;

    static void Main(string[] args)
    {
        string pdfPath = args.Length > 0 ? args[0] : GetPdfPathFromUser();

        var processors = new List<IColumnProcessor>
        {
            new TaskProcessor(),
            new TagProcessor(),
            new DurationProcessor(),
            new TimeProcessor()
        };

        var extractor = new PdfDataExtractor(processors);
        List<PdfData> extractedData = extractor.Extract(pdfPath);

        // Exploiter les données pour le regroupement
        var weeklyTagSummaries = extractedData
            .GroupBy(data => new
            {
                WeekNumber = GetWeekNumber(data.StartTime),
                data.Tag
            })
            .Select(g => new
            {
                WeekNumber = g.Key.WeekNumber,
                Tag = g.Key.Tag,
                TotalHours = g.Sum(x => x.Duration.TotalHours)
            })
            .ToList();


        // Affichage des résultats
        Console.WriteLine("");
        foreach (var summary in weeklyTagSummaries.OrderBy(s => s.WeekNumber))
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write($"S{summary.WeekNumber} - ");

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.Write($"{summary.Tag}: ");

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"{summary.TotalHours}h");

            Console.ResetColor();
        }

        // Totaux hebdomadaires
        var weeklyTotals = weeklyTagSummaries
            .GroupBy(summary => summary.WeekNumber)
            .Select(g => new
            {
                WeekNumber = g.Key,
                TotalHours = g.Sum(x => x.TotalHours)
            })
            .ToList();

        Console.WriteLine("");
        foreach (var total in weeklyTotals)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write($"S{total.WeekNumber} - ");

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"{total.TotalHours}h");

            Console.ResetColor();
        }

        // Totaux période
        var periodTotals = weeklyTagSummaries
            .GroupBy(summary => summary.Tag)
            .Select(g => new
            {
                WeekNumber = g.Key,
                TotalHours = g.Sum(x => x.TotalHours)
            })
            .ToList();

        Console.WriteLine("");
        foreach (var total in periodTotals)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write($"{total.WeekNumber} - ");

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"{total.TotalHours}h");

            Console.ResetColor();
        }

        Console.WriteLine("");
        ExcelGenerator generator = new ExcelGenerator(extractedData, HOURS_PER_DAY_WORKED, HOURS_PER_DAY_BILLABLE, pdfPath.Replace(".pdf", ".xlsx"));
        generator.CreateExcel();
        Console.WriteLine("Generated excel");
    }

    private static int GetWeekNumber(DateTime date)
    {
        var calendar = CultureInfo.CurrentCulture.Calendar;
        var weekOfYear = calendar.GetWeekOfYear(date, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
        return weekOfYear;
    }

    private static string GetPdfPathFromUser()
    {
        Console.WriteLine("Veuillez glisser-déposer un rapport détaillé Toggl dans la console:");
        return Console.ReadLine()?.Trim('\"');
    }
}
