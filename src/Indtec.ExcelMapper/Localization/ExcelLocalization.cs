namespace Indtec.ExcelMapper.Localization;

public enum ExcelLanguage
{
    English,
    PortugueseBrazil
}

public sealed class ExcelMapperOptions
{
    public ExcelLanguage Language { get; set; } = ExcelLanguage.English;
    public IExcelMessageProvider? Messages { get; set; }

    internal IExcelMessageProvider ResolveMessages()
        => Messages ?? (Language == ExcelLanguage.PortugueseBrazil
            ? ExcelMessages.PortugueseBrazil
            : ExcelMessages.English);
}

public interface IExcelMessageProvider
{
    string BatchValidatorsRequireAsync();
    string WorksheetNotFound(string sheetName);
    string RequiredColumnsNotFound(string sheetName, IReadOnlyList<string> columns);
    string ReadOnlyProperty(string header);
    string AtLeastOneSheetForImport();
    string DuplicateWorkbookModel(string modelName);
    string AtLeastOneSheetForTemplate();
    string DuplicateWorksheet(string sheetName);
    string InvalidTemplateRows();
    string AllowedValuesTooLong(string header);
    string InvalidAllowedValue(string header);
}

public static class ExcelMessages
{
    public static IExcelMessageProvider English { get; } = new EnglishExcelMessageProvider();
    public static IExcelMessageProvider PortugueseBrazil { get; } = new PortugueseBrazilExcelMessageProvider();
}

internal sealed class EnglishExcelMessageProvider : IExcelMessageProvider
{
    public string BatchValidatorsRequireAsync() => "Batch validators require ImportAsync because they may perform asynchronous work.";
    public string WorksheetNotFound(string sheetName) => $"Worksheet '{sheetName}' was not found.";
    public string RequiredColumnsNotFound(string sheetName, IReadOnlyList<string> columns) => $"Required columns were not found in worksheet '{sheetName}': {string.Join(", ", columns)}.";
    public string ReadOnlyProperty(string header) => $"Property mapped to '{header}' is read-only and cannot be imported.";
    public string AtLeastOneSheetForImport() => "At least one sheet must be registered for workbook import.";
    public string DuplicateWorkbookModel(string modelName) => $"Model '{modelName}' was registered more than once in the workbook import.";
    public string AtLeastOneSheetForTemplate() => "At least one sheet must be registered for workbook template generation.";
    public string DuplicateWorksheet(string sheetName) => $"Worksheet '{sheetName}' was registered more than once.";
    public string InvalidTemplateRows() => "TemplateRows must be greater than zero.";
    public string AllowedValuesTooLong(string header) => $"Allowed values for '{header}' exceed Excel's 255-character inline validation limit.";
    public string InvalidAllowedValue(string header) => $"Choose one of the allowed values for {header}.";
}

internal sealed class PortugueseBrazilExcelMessageProvider : IExcelMessageProvider
{
    public string BatchValidatorsRequireAsync() => "Validadores em lote exigem ImportAsync porque podem executar operações assíncronas.";
    public string WorksheetNotFound(string sheetName) => $"A planilha '{sheetName}' não foi encontrada.";
    public string RequiredColumnsNotFound(string sheetName, IReadOnlyList<string> columns) => $"Colunas obrigatórias não foram encontradas na planilha '{sheetName}': {string.Join(", ", columns)}.";
    public string ReadOnlyProperty(string header) => $"A propriedade mapeada para '{header}' é somente leitura e não pode ser importada.";
    public string AtLeastOneSheetForImport() => "Ao menos uma planilha deve ser registrada para importar o workbook.";
    public string DuplicateWorkbookModel(string modelName) => $"O modelo '{modelName}' foi registrado mais de uma vez na importação do workbook.";
    public string AtLeastOneSheetForTemplate() => "Ao menos uma planilha deve ser registrada para gerar o template do workbook.";
    public string DuplicateWorksheet(string sheetName) => $"A planilha '{sheetName}' foi registrada mais de uma vez.";
    public string InvalidTemplateRows() => "TemplateRows deve ser maior que zero.";
    public string AllowedValuesTooLong(string header) => $"Os valores permitidos para '{header}' excedem o limite de 255 caracteres da validação inline do Excel.";
    public string InvalidAllowedValue(string header) => $"Escolha um dos valores permitidos para {header}.";
}
