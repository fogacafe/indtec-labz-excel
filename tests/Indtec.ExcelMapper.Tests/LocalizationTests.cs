using ClosedXML.Excel;
using Indtec.ExcelMapper.Localization;
using Xunit;

namespace Indtec.ExcelMapper.Tests;

public sealed class LocalizationTests
{
    [Fact]
    public void Import_ShouldUsePortugueseBrazilProcessingMessages()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Products");
            sheet.Cell(1, 1).Value = "Id";
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var mapper = new ExcelMapper(options =>
            options.Language = ExcelLanguage.PortugueseBrazil);

        var error = Assert.Throws<ExcelMappingException>(() => mapper.Import<ProductRow>(stream));

        Assert.Contains("Colunas obrigatórias", error.Message);
        Assert.Contains("Products", error.Message);
    }

    [Fact]
    public void Import_ShouldAllowMessageProviderOverride()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.AddWorksheet("Other");
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var mapper = new ExcelMapper(options =>
            options.Messages = new CustomMessages());

        var error = Assert.Throws<ExcelMappingException>(() => mapper.Import<ProductRow>(stream));

        Assert.Equal("CUSTOM: Products", error.Message);
    }
}

internal sealed class CustomMessages : ExcelMessageProvider
{
    public override string WorksheetNotFound(string sheetName)
        => $"CUSTOM: {sheetName}";
}
