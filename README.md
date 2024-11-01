## Description

The **WordReporter.NPOI** NuGet package provides an easy way to generate Word documents based on templates using the NPOI library. It allows you to quickly and effortlessly fill templates with data and save them in `.docx` format.

## Installation

You can install the package via the NuGet Package Manager by running the following command:

```
Install-Package WordReporter.NPOI
```

Or using the .NET CLI:

```
dotnet add package WordReporter.NPOI
```

## Usage

### Code Example

Here’s an example of how to use the package to generate a Word document based on a template:

```csharp
// Path to the template and resulting file
string templatePath = "path_to_template.docx";
string resultPath = "path_to_result.docx";

// Data to be filled into the template
var data = new
{
    // Your data
};

// Create an instance of XWPFTemplate with the specified template
var template = new XWPFTemplate(templatePath);

// Generate the document using the provided data
template.Generate(data);

// Save the generated document to a file
using var t = new FileStream(resultPath, FileMode.Create, FileAccess.Write);
template.SaveAs(t);
```

### Parameter Descriptions

- `templatePath`: The path to the template file, which should be in `.docx` format.
- `resultPath`: The path to the file where the generated document will be saved.
- `data`: An object containing the data to fill the template. Ensure that the fields of the object correspond to the placeholders in the template.

## Dependencies

- .NET 8.0.
- NPOI

## License

This project is licensed under the Apache License, Version 2.0. You can find the full license text in the [LICENSE](LICENSE) file.

## Additional Information
