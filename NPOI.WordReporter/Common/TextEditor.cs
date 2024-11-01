using System.Reflection;
using Scriban;

namespace NPOI.WordReporter.Common;

public static class TextEditor
{
    public static string ReplaceText(string text, object item) //TODO: Remove Scriban or use everywhere
    {
        var template = Template.Parse(text);
        var result = template.Render(item, memberRenamer: member => member.Name);
        return result;
    }

    public static string ReplacePlaceholders(string text, object item)
    {
        var itemType = item.GetType();

        foreach (var prop in itemType.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            var placeholder = $"{{{{item.{prop.Name}}}}}";

            if (!text.Contains(placeholder))
                continue;

            var value = prop.GetValue(item)?.ToString() ?? string.Empty;
            text = text.Replace(placeholder, value);
        }

        return text;
    }
}