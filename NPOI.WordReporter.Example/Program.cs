using NPOI.WordReporter;
using NPOI.WordReporter.Example;

var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.docx");
var resultPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "result.docx");

var data = new Dto
{
    CustomerName = "It is customer",
    Main = new Main
    {
        Description = "It is long description",
        Title = "Important document"
    },
    Table1 =
    [
        new Table1 { Column1 = "Cell value 1", Column2 = "Cell value 2" },
        new Table1 { Column1 = "Cell value 3", Column2 = "Cell value 4" }
    ]
};

// Load the template
var template = new XWPFTemplate(templatePath);
template.Generate(data);

// Save the result document (overwriting if it exists)
using var t = new FileStream(resultPath, FileMode.Create, FileAccess.Write);
template.SaveAs(t);