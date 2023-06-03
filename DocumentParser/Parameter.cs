namespace DocumentParser;

public class Parameter
{
    public string? Signal { get; set; }
    public string? Designation { get; set; }
    public string? TypeSignal { get; set; }
    public string? Unit { get; set; }
    public List<double>? ChangeRange { get; set; }
    public string? Address { get; set; }
    public string? HighDischargesPrice { get; set; }
    public string? QuantityMeaningDischarges{ get; set; }
    public  string? /*int?*/ FrequencyRegister { get; set; }
}