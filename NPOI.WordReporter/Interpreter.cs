using NPOI.WordReporter.Elements;
using NPOI.XWPF.UserModel;

namespace NPOI.WordReporter;

internal class Interpreter
{
    public void Evaluate(XWPFDocument? document, object dto)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));

        //Отказаться от Scriban, он нужен лишь чтобы заменить значение из DTO
        foreach (var element in document.BodyElements)
        {
            var handler = DocumentElementHandlerFactory.GetHandler(element);
            handler?.Handle(dto);
        }
    }
}