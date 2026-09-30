namespace Indtec.ExcelMapper;

/// <summary>Represents an error produced while mapping, validating or generating an Excel workbook.</summary>
public class ExcelMappingException : Exception
{
    /// <summary>Creates a mapping exception with the supplied message.</summary>
    public ExcelMappingException(string message) : base(message) { }

    /// <summary>Creates a mapping exception with the supplied message and inner exception.</summary>
    public ExcelMappingException(string message, Exception innerException) : base(message, innerException) { }
}

/// <summary>Thrown when a model does not expose the source-generated Excel map required by the mapper.</summary>
public sealed class GeneratedMapNotFoundException : ExcelMappingException
{
    /// <summary>Creates an exception for a model type without a generated Excel map.</summary>
    public GeneratedMapNotFoundException(Type type)
        : base($"No generated Excel map was found for '{type.FullName}'. Mark the class as partial and annotate it with [ExcelSheet].")
    {
    }
}
