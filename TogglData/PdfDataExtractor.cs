using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;

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
