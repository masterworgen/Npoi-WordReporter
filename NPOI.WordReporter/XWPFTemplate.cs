using NPOI.XWPF.UserModel;

namespace NPOI.WordReporter;

// ReSharper disable once InconsistentNaming
public class XWPFTemplate(string fileName) : IDisposable
{
    private Stream? _fileStream;
    private readonly Interpreter _interpreter = new();
    private XWPFDocument? _document;

    public void Generate(object data)
    {
        if (!File.Exists(fileName))
            throw new FileNotFoundException("Not found file", fileName);

        _fileStream = new FileStream(fileName, FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite);
        _fileStream.Position = 0;

        _document = new XWPFDocument(_fileStream);

        _interpreter.Evaluate(_document, data);
    }

    public void SaveAs(Stream stream)
    {
        if (_document is null)
            throw new NullReferenceException("Document has not been generated. Please call Generate() before SaveAs().");

        _document.Write(stream);
        stream.Seek(0, SeekOrigin.Begin);
    }

    public void Dispose()
    {
        _fileStream?.Dispose();
        _document?.Dispose();
    }
}