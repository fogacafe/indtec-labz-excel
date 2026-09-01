using ClosedXML.Excel;
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
        Assert.Equal(1, result[0].Id);
        Assert.Equal("Coffee", result[0].Name);
        Assert.Equal(10m, result[0].Cost);
        Assert.Equal(12.50m, result[0].Price);
        Assert.True(result[0].Active);
        Assert.Equal(2, result[1].Id);
        Assert.Equal("Tea", result[1].Name);
        Assert.Equal(9m, result[1].Cost);
        Assert.Equal(8.75m, result[1].Price);
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
            options.Header.Bold().Background("#1F2937").FontColor("#FFFFFF");
            options.Column(x => x.Price)
                .NumberFormat("#,##0.00")
                .Width(18)
                .When(row => row.Price < row.Cost)
                .Background("#FFCCCC")
                .Bold();
            options.Row()
                .When(row => !row.Active)
                .FontColor("#999999");
        });

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        var sheet = workbook.Worksheet("Products");

        Assert.True(sheet.Cell(1, 1).Style.Font.Bold);
        Assert.Equal("#,##0.00", sheet.Cell(2, 4).Style.NumberFormat.Format);
        Assert.Equal(18d, sheet.Column(4).Width);
        Assert.Equal(XLColor.FromHtml("#FFCCCC"), sheet.Cell(3, 4).Style.Fill.BackgroundColor);
        Assert.True(sheet.Cell(3, 4).Style.Font.Bold);
        Assert.Equal(XLColor.FromHtml("#999999"), sheet.Cell(3, 2).Style.Font.FontColor);
    }
}

[ExcelSheet("Products")]
public partial class ProductRow
{
    [ExcelColumn("Id", Order = 1)]
    public int Id { get; set; }

    [ExcelColumn("Name", Order = 2)]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("Cost", Order = 3)]
    public decimal Cost { get; set; }

    [ExcelColumn("Price", Order = 4)]
    public decimal Price { get; set; }

    [ExcelColumn("Active", Order = 5)]
    public bool Active { get; set; }
}
