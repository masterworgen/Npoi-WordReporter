using System.Reflection;

namespace NPOI.WordReporter.Common;

public static class PropertyPathResolver
{
    public static object? GetPropertyValue(object obj, string propertyPath)
    {
        if (obj == null || string.IsNullOrWhiteSpace(propertyPath))
            throw new ArgumentException("Object or property path cannot be null or empty.");

        var properties = propertyPath.Split('.');
        var currentObject = obj;

        foreach (var propertyName in properties)
        {
            if (currentObject == null)
                return null;

            var propertyInfo = currentObject.GetType().GetProperty(propertyName,
                BindingFlags.Instance | BindingFlags.Public);

            if (propertyInfo == null)
                throw new ArgumentException($"Property '{propertyName}' not found in type '{currentObject.GetType().Name}'");

            currentObject = propertyInfo.GetValue(currentObject);
        }

        return currentObject;
    }
}