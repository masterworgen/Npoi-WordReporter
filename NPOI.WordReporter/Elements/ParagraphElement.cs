using NPOI.WordReporter.Common;
using NPOI.WordReporter.Elements.Abstracts;
using NPOI.XWPF.UserModel;
using System.Text.RegularExpressions;

namespace NPOI.WordReporter.Elements;

internal class ParagraphElement(XWPFParagraph element) : DocumentElementHandler(element)
{
    private readonly Regex _variableRegex = new(@"{{\s*(.+?)\s*}}");

    public override void Handle(object dto)
    {
        if (string.IsNullOrWhiteSpace(element.Text))
            return;

        ReplaceText(dto);
    }

    private void ReplaceText(object dto)
    {
        var matches = _variableRegex.Matches(element.Text);

        foreach (var match in matches)
        {
            var variable = match.ToString()!;
            var textReplaced = TextEditor.ReplaceText(variable, dto);
            element.ReplaceText(variable, textReplaced);
        }
    }
}
