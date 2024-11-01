using NPOI.XWPF.UserModel;

namespace NPOI.WordReporter.Elements.Abstracts;

internal abstract class DocumentElementHandler(IBodyElement element)
{
    public abstract void Handle(object dto);
}