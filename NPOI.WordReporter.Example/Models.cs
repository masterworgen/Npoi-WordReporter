namespace NPOI.WordReporter.Example;

public class Dto
{
    public string CustomerName { get; set; } = string.Empty;
    public Main Main { get; set; } = new();
    public Table1[] Table1 { get; set; } = [];
}

public class Main
{
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}

public class Table1
{
    public string Column1 { get; set; } = string.Empty;
    public string Column2 { get; set; } = string.Empty;
}