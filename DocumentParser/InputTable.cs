namespace DocumentParser;

public class InputTable
{
    public string Name { get; set; }

    public IList<Parameter> Parameters { get; set; } = new List<Parameter>();
}