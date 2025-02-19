
using System.Globalization;
using SpreadsheetLight;

public class ExcelGenerator
{
    public List<PdfData> Data { get; }
    public string Filepath { get; }
    public double TimeBase { get; }

    public ExcelGenerator(List<PdfData> data, double timeBase, string filepath)
    {
        TimeBase = timeBase;
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
        doc.SetColumnWidth(column++, 15);
        doc.SetColumnWidth(column++, 40);

        column = 1;

        doc.SetCellValue(1, column++, "Projet");
        doc.SetCellValue(1, column++, "Description");
        doc.SetCellValue(1, column++, "Début");
        doc.SetCellValue(1, column++, "Fin");
        doc.SetCellValue(1, column++, "Semaine");
        doc.SetCellValue(1, column++, "RMA");
        doc.SetCellValue(1, column++, "Mois");
        doc.SetCellValue(1, column++, "Durée (h)");
        doc.SetCellValue(1, column++, $"Journées (/{TimeBase}h)");
        doc.SetCellValue(1, column++, "Tags");

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

            double days = Math.Round((hours / TimeBase) , 2);
            string daysString = days.ToString().Replace(',', '.');
            doc.SetCellValueNumeric(row, column++, daysString);

            string[] tags = line.Tag.Split(',');
            foreach(string tag in tags.Reverse())
            {
                doc.SetCellValue(row, column++, tag);
            }
        }

        doc.SaveAs(Filepath);
    }
}
