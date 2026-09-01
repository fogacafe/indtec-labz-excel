namespace Indtec.ExcelMapper;

public class ExcelMappingException : Exception
{
    public ExcelMappingException(string message) : base(message) { }
    public ExcelMappingException(string message, Exception innerException) : base(message, innerException) { }
}

public sealed class GeneratedMapNotFoundException : ExcelMappingException
{
    public GeneratedMapNotFoundException(Type type)
        : base($"No generated Excel map was found for '{type.FullName}'. Mark the class as partial and annotate it with [ExcelSheet].")
    {
    }
}
