using ClosedXML.Excel;
using Indtec.ExcelMapper.Importing;
using Indtec.ExcelMapper.Localization;
using Xunit;

namespace Indtec.ExcelMapper.Tests;

public sealed class StreamingChunkImportTests
{
    [Fact]
    public async Task ImportChunksAsync_ShouldDeliverBoundedChunksWithoutAccumulatingRows()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("StreamingRows");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(1, 3).Value = "Amount";
            sheet.Cell(1, 4).Value = "TradeDate";

            for (var i = 1; i <= 2500; i++)
            {
                var row = i + 1;
                sheet.Cell(row, 1).Value = i;
                sheet.Cell(row, 2).Value = $"Item {i}";
                sheet.Cell(row, 3).Value = i + 0.25m;
                sheet.Cell(row, 4).Value = new DateTime(2026, 9, 1).AddDays(i % 20);
            }

            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var mapper = new ExcelMapper();
        var chunks = new List<(int Count, int Start, int End)>();
        var totalItems = 0;

        await mapper.ImportChunksAsync<StreamingRow>(
            stream,
            (chunk, _) =>
            {
                Assert.InRange(chunk.Rows.Count, 1, 1000);
                Assert.Empty(chunk.Errors);
                totalItems += chunk.Items.Count;
                chunks.Add((chunk.Rows.Count, chunk.StartRow, chunk.EndRow));
                return Task.CompletedTask;
            },
            chunkSize: 1000);

        Assert.Equal(2500, totalItems);
        Assert.Equal(3, chunks.Count);
        Assert.Equal((1000, 2, 1001), chunks[0]);
        Assert.Equal((1000, 1002, 2001), chunks[1]);
        Assert.Equal((500, 2002, 2501), chunks[2]);
    }

    [Fact]
    public async Task ImportChunksAsync_ShouldCollectLocalValidationAndConversionErrors()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("StreamingRows");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(1, 3).Value = "Amount";
            sheet.Cell(1, 4).Value = "TradeDate";

            sheet.Cell(2, 1).Value = 1;
            sheet.Cell(2, 2).Value = "Valid";
            sheet.Cell(2, 3).Value = 10;
            sheet.Cell(2, 4).Value = new DateTime(2026, 9, 1);

            sheet.Cell(3, 1).Value = 2;
            sheet.Cell(3, 2).Value = "Invalid amount";
            sheet.Cell(3, 3).Value = -1;
            sheet.Cell(3, 4).Value = new DateTime(2026, 9, 1);

            sheet.Cell(4, 1).Value = "bad-id";
            sheet.Cell(4, 2).Value = "Invalid id";
            sheet.Cell(4, 3).Value = 10;
            sheet.Cell(4, 4).Value = new DateTime(2026, 9, 1);

            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var mapper = new ExcelMapper(options =>
            options.Language = ExcelLanguage.PortugueseBrazil);
        var errors = new List<ExcelImportError>();
        var validItems = 0;

        await mapper.ImportChunksAsync<StreamingRow>(
            stream,
            (chunk, _) =>
            {
                errors.AddRange(chunk.Errors);
                validItems += chunk.Items.Count;
                return Task.CompletedTask;
            },
            chunkSize: 2,
            configure: options =>
            {
                options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
                options.Validate(x => x.Amount >= 0, "Amount must be positive.");
            });

        Assert.Equal(1, validItems);
        Assert.Equal(2, errors.Count);
        Assert.Contains(errors, x => x.Row == 3 && x.Message == "Amount must be positive.");
        Assert.Contains(errors, x => x.Row == 4 && x.Message.Contains("Não foi possível converter"));
    }

    [Fact]
    public async Task ImportChunksAsync_ShouldRejectFullSheetBatchValidators()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("StreamingRows");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(1, 3).Value = "Amount";
            sheet.Cell(1, 4).Value = "TradeDate";
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var mapper = new ExcelMapper();

        var error = await Assert.ThrowsAsync<ExcelMappingException>(() =>
            mapper.ImportChunksAsync<StreamingRow>(
                stream,
                (_, _) => Task.CompletedTask,
                configure: options => options.AddBatchValidator(new StreamingBatchValidator())));

        Assert.Contains("full-sheet context", error.Message);
    }
}

internal sealed class StreamingBatchValidator : IExcelBatchValidator<StreamingRow>
{
    public Task<IReadOnlyList<ExcelImportError>> ValidateAsync(
        ExcelBatchValidationContext<StreamingRow> context,
        CancellationToken cancellationToken = default)
        => Task.FromResult<IReadOnlyList<ExcelImportError>>(Array.Empty<ExcelImportError>());
}

[ExcelSheet("StreamingRows")]
public partial class StreamingRow
{
    [ExcelColumn("Id", Order = 1, Required = true)]
    public int Id { get; set; }

    [ExcelColumn("Name", Order = 2, Required = true)]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("Amount", Order = 3, Required = true)]
    public decimal Amount { get; set; }

    [ExcelColumn("TradeDate", Order = 4, Required = true)]
    public DateTime TradeDate { get; set; }
}
