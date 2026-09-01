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
    string StreamingBatchValidatorsNotSupported();
    string WorksheetNotFound(string sheetName);
    string RequiredColumnsNotFound(string sheetName, IReadOnlyList<string> columns);
    string ReadOnlyProperty(string header);
    string EmptyCellNotNullable(string address, string typeName);
    string CouldNotConvertCell(string address, string typeName);
    string MissingCellReference();
    string InvalidCellReference(string reference);
    string AtLeastOneSheetForImport();
    string DuplicateWorkbookModel(string modelName);
    string UnregisteredWorkbookModel(string modelName);
    string AtLeastOneSheetForTemplate();
    string DuplicateWorksheet(string sheetName);
    string InvalidTemplateRows();
    string AllowedValuesTooLong(string header);
    string InvalidValueTitle();
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
    public string StreamingBatchValidatorsNotSupported() => "Batch validators are not supported by streaming chunk imports because they require full-sheet context. Validate inside the chunk callback or use ImportAsync.";
    public string WorksheetNotFound(string sheetName) => $"Worksheet '{sheetName}' was not found.";
    public string RequiredColumnsNotFound(string sheetName, IReadOnlyList<string> columns) => $"Required columns were not found in worksheet '{sheetName}': {string.Join(", ", columns)}.";
    public string ReadOnlyProperty(string header) => $"Property mapped to '{header}' is read-only and cannot be imported.";
    public string EmptyCellNotNullable(string address, string typeName) => $"Cell {address} is empty but '{typeName}' is not nullable.";
    public string CouldNotConvertCell(string address, string typeName) => $"Could not convert cell {address} to '{typeName}'.";
    public string MissingCellReference() => "A streamed cell is missing its Excel reference.";
    public string InvalidCellReference(string reference) => $"Invalid Excel cell reference '{reference}'.";
    public string AtLeastOneSheetForImport() => "At least one sheet must be registered for workbook import.";
    public string DuplicateWorkbookModel(string modelName) => $"Model '{modelName}' was registered more than once in the workbook import.";
    public string UnregisteredWorkbookModel(string modelName) => $"Workbook validator returned an error for unregistered model '{modelName}'.";
    public string AtLeastOneSheetForTemplate() => "At least one sheet must be registered for workbook template generation.";
    public string DuplicateWorksheet(string sheetName) => $"Worksheet '{sheetName}' was registered more than once.";
    public string InvalidTemplateRows() => "TemplateRows must be greater than zero.";
    public string AllowedValuesTooLong(string header) => $"Allowed values for '{header}' exceed Excel's 255-character inline validation limit.";
    public string InvalidValueTitle() => "Invalid value";
    public string InvalidAllowedValue(string header) => $"Choose one of the allowed values for {header}.";
}

internal sealed class PortugueseBrazilExcelMessageProvider : IExcelMessageProvider
{
    public string BatchValidatorsRequireAsync() => "Validadores em lote exigem ImportAsync porque podem executar operações assíncronas.";
    public string StreamingBatchValidatorsNotSupported() => "Validadores em lote não são suportados na importação streaming em chunks porque exigem o contexto completo da planilha. Valide no callback do chunk ou use ImportAsync.";
    public string WorksheetNotFound(string sheetName) => $"A planilha '{sheetName}' não foi encontrada.";
    public string RequiredColumnsNotFound(string sheetName, IReadOnlyList<string> columns) => $"Colunas obrigatórias não foram encontradas na planilha '{sheetName}': {string.Join(", ", columns)}.";
    public string ReadOnlyProperty(string header) => $"A propriedade mapeada para '{header}' é somente leitura e não pode ser importada.";
    public string EmptyCellNotNullable(string address, string typeName) => $"A célula {address} está vazia, mas '{typeName}' não aceita nulo.";
    public string CouldNotConvertCell(string address, string typeName) => $"Não foi possível converter a célula {address} para '{typeName}'.";
    public string MissingCellReference() => "Uma célula lida em streaming não possui referência do Excel.";
    public string InvalidCellReference(string reference) => $"A referência de célula do Excel '{reference}' é inválida.";
    public string AtLeastOneSheetForImport() => "Ao menos uma planilha deve ser registrada para importar o workbook.";
    public string DuplicateWorkbookModel(string modelName) => $"O modelo '{modelName}' foi registrado mais de uma vez na importação do workbook.";
    public string UnregisteredWorkbookModel(string modelName) => $"O validador do workbook retornou um erro para o modelo não registrado '{modelName}'.";
    public string AtLeastOneSheetForTemplate() => "Ao menos uma planilha deve ser registrada para gerar o template do workbook.";
    public string DuplicateWorksheet(string sheetName) => $"A planilha '{sheetName}' foi registrada mais de uma vez.";
    public string InvalidTemplateRows() => "TemplateRows deve ser maior que zero.";
    public string AllowedValuesTooLong(string header) => $"Os valores permitidos para '{header}' excedem o limite de 255 caracteres da validação inline do Excel.";
    public string InvalidValueTitle() => "Valor inválido";
    public string InvalidAllowedValue(string header) => $"Escolha um dos valores permitidos para {header}.";
}
