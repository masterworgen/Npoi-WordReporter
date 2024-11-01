using System.Collections;
using NPOI.WordReporter.Common;
using NPOI.WordReporter.Elements.Abstracts;
using NPOI.XWPF.UserModel;

namespace NPOI.WordReporter.Elements;

internal class TableElement(XWPFTable element) : DocumentElementHandler(element)
{
    public override void Handle(object dto)
    {
        var bookmark = GetBookmarkInTable();

        if (string.IsNullOrWhiteSpace(bookmark))
        {
            // Static
            foreach (var row in element.Rows)
            {
                foreach (var cell in row.GetTableCells())
                {
                    foreach (var paragraph in cell.Paragraphs)
                    {
                        var handler = DocumentElementHandlerFactory.GetHandler(paragraph);
                        handler!.Handle(dto);
                    }
                }
            }
        }
        else
        {
            // Dynamic
            var items = (IEnumerable?)PropertyPathResolver.GetPropertyValue(dto, bookmark);
            FillDynamicTable(items ?? Array.Empty<object>());
        }
    }

    private string? GetBookmarkInTable()
    {
        foreach (var row in element.Rows)
        {
            foreach (var cell in row.GetTableCells())
            {
                foreach (var paragraph in cell.Paragraphs)
                {
                    foreach (var item in paragraph.GetCTP().GetBookmarkStartList())
                    {
                        return item.name;
                    }
                }
            }
        }

        return null;

    }

    private void FillDynamicTable(IEnumerable items)
    {
        var templateRow = element.GetRow(1);

        while (element.NumberOfRows > 2)
        {
            element.RemoveRow(2);
        }

        foreach (var item in items)
        {
            var newRow = templateRow.CloneRow();

            foreach (var xwpfTableCell in newRow.GetTableCells())
            {
                foreach (var xwpfParagraph in xwpfTableCell.Paragraphs)
                {
                    var text = xwpfParagraph.Text;
                    var replacedText = TextEditor.ReplacePlaceholders(text, item);
                    xwpfParagraph.ReplaceText(text, replacedText);
                }
            }
        }

        element.RemoveRow(1);
    }
}
