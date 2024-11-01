using NPOI.WordReporter.Elements.Abstracts;
using NPOI.XWPF.UserModel;

namespace NPOI.WordReporter.Elements;

internal static class DocumentElementHandlerFactory
{
    public static DocumentElementHandler? GetHandler(IBodyElement element) =>
        element switch
        {
            XWPFParagraph paragraph => new ParagraphElement(paragraph),
            XWPFTable table => new TableElement(table),
            _ => null
        };
}