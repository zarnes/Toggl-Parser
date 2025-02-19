
using System.Globalization;

public class Program
{
    static void Main(string[] args)
    {
        string pdfPath = args.Length >= 1 ? args[0] : GetPdfPathFromUser();
        int hoursPerDay = args.Length >= 2 ? int.Parse(args[1]) : GetHoursPerDayFromUser();

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
        ExcelGenerator generator = new ExcelGenerator(extractedData, hoursPerDay, pdfPath.Replace(".pdf", ".xlsx"));
        generator.CreateExcel();
        Console.WriteLine("Generated excel");
    }

    private static int GetHoursPerDayFromUser()
    {
        int hoursPerDay;
        Console.WriteLine("Veuillez entrer le nombre d'heures par jour : ");
        while (!int.TryParse(Console.ReadLine(), out hoursPerDay))
        {
            Console.WriteLine("Entrée invalide. Veuillez entrer un nombre valide d'heures par jour : ");
        }
        return hoursPerDay;
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
