public interface IColumnProcessor
{
    void BufferText(string text);
    void Process(PdfData data);
    void Reset();
}
