using ClosedXML.Excel;
using Indtec.ExcelMapper.Conversion;
using Indtec.ExcelMapper.Styling;
using Xunit;

namespace Indtec.ExcelMapper.Tests;

public sealed class ExcelMapperTests
{
    [Fact]
    public void ExportThenImport_ShouldPreserveMappedValues()
    {
        var mapper = new ExcelMapper();
        var source = new[]
        {
            new ProductRow { Id = 1, Name = "Coffee", Cost = 10m, Price = 12.50m, Active = true },
            new ProductRow { Id = 2, Name = "Tea", Cost = 9m, Price = 8.75m, Active = false }
        };

        using var stream = new MemoryStream();
        mapper.Export(source, stream);
        stream.Position = 0;

        var result = mapper.Import<ProductRow>(stream);

        Assert.Equal(2, result.Count);
        Assert.Equal("Coffee", result[0].Name);
        Assert.True(result[0].Active);
        Assert.Equal("Tea", result[1].Name);
        Assert.False(result[1].Active);
    }

    [Fact]
    public void Export_WhenRuleComparesOtherCell_ShouldStyleTargetCell()
    {
        var mapper = new ExcelMapper();
        var source = new[]
        {
            new ProductRow { Id = 1, Name = "Coffee", Cost = 10m, Price = 12m, Active = true },
            new ProductRow { Id = 2, Name = "Tea", Cost = 10m, Price = 8m, Active = false }
        };

        using var stream = new MemoryStream();
        mapper.Export(source, stream, options =>
        {
            options.UseTheme(new ProductTheme());
            options.Column(x => x.Price)
                .When(row => row.Price < row.Cost)
                .Background("#FFCCCC")
                .Bold();
        });

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        var sheet = workbook.Worksheet("Products");

        Assert.True(sheet.Cell(1, 1).Style.Font.Bold);
        Assert.Equal("#,##0.00", sheet.Cell(2, 4).Style.NumberFormat.Format);
        Assert.Equal(18d, sheet.Column(4).Width);
        Assert.Equal(XLColor.FromHtml("#FFCCCC"), sheet.Cell(3, 4).Style.Fill.BackgroundColor);
        Assert.Equal("Yes", sheet.Cell(2, 5).GetString());
        Assert.Equal("No", sheet.Cell(3, 5).GetString());
    }

    [Fact]
    public void Import_WhenRequiredColumnIsMissing_ShouldFailBeforeMappingRows()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Products");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(2, 1).Value = 1;
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var mapper = new ExcelMapper();

        var error = Assert.Throws<ExcelMappingException>(() => mapper.Import<ProductRow>(stream));
        Assert.Contains("Name", error.Message);
    }
}

public sealed class ProductTheme : ExcelTheme<ProductRow>
{
    public override void Configure(ExcelExportOptions<ProductRow> options)
    {
        options.Header.Bold().Background("#1F2937").FontColor("#FFFFFF");
        options.Column(x => x.Price).NumberFormat("#,##0.00").Width(18);
        options.Row().When(row => !row.Active).FontColor("#999999");
    }
}

public sealed class YesNoBoolConverter : IExcelValueConverter
{
    public object? Read(ExcelValue value, Type destinationType)
    {
        if (value.IsEmpty) return false;
        return string.Equals(value.AsString(), "Yes", StringComparison.OrdinalIgnoreCase);
    }

    public ExcelValue Write(object? value)
        => new(value is true ? "Yes" : "No");
}

[ExcelSheet("Products")]
public partial class ProductRow
{
    [ExcelColumn("Id", Order = 1, Required = true)]
    public int Id { get; set; }

    [ExcelColumn("Name", Order = 2, Required = true)]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("Cost", Order = 3)]
    public decimal Cost { get; set; }

    [ExcelColumn("Price", Order = 4)]
    public decimal Price { get; set; }

    [ExcelColumn("Active", Order = 5, Converter = typeof(YesNoBoolConverter))]
    public bool Active { get; set; }
}
