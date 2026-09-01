using ClosedXML.Excel;
using Indtec.ExcelMapper.Conversion;
using Indtec.ExcelMapper.Importing;
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

    [Fact]
    public void Import_CollectMode_ShouldReturnValidItemsAndRowErrors()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Products");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(1, 3).Value = "Cost";
            sheet.Cell(1, 4).Value = "Price";
            sheet.Cell(1, 5).Value = "Active";
            sheet.Cell(1, 6).Value = "Status";

            sheet.Cell(2, 1).Value = 1;
            sheet.Cell(2, 2).Value = "Coffee";
            sheet.Cell(2, 3).Value = 10;
            sheet.Cell(2, 4).Value = 12;
            sheet.Cell(2, 5).Value = "Yes";
            sheet.Cell(2, 6).Value = "Active";

            sheet.Cell(3, 1).Value = 2;
            sheet.Cell(3, 2).Value = "Tea";
            sheet.Cell(3, 3).Value = 10;
            sheet.Cell(3, 4).Value = "not-a-price";
            sheet.Cell(3, 5).Value = "No";
            sheet.Cell(3, 6).Value = "Inactive";

            sheet.Cell(4, 1).Value = 3;
            sheet.Cell(4, 2).Value = "Milk";
            sheet.Cell(4, 3).Value = 10;
            sheet.Cell(4, 4).Value = 5;
            sheet.Cell(4, 5).Value = "Yes";
            sheet.Cell(4, 6).Value = "Active";
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var mapper = new ExcelMapper();
        var result = mapper.Import<ProductRow>(stream, options =>
        {
            options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
            options.Validate(row => row.Price >= row.Cost, "Price cannot be lower than cost.");
        });

        Assert.Single(result.Items);
        Assert.Equal("Coffee", result.Items[0].Name);
        Assert.Equal(2, result.Errors.Count);
        Assert.Contains(result.Errors, error => error.Row == 3 && error.Column == "Price");
        Assert.Contains(result.Errors, error => error.Row == 4 && error.Message.Contains("lower than cost"));
    }

    [Fact]
    public void CreateTemplate_ShouldGenerateHeadersStylesAndDropdowns()
    {
        var mapper = new ExcelMapper();
        using var stream = new MemoryStream();

        mapper.CreateTemplate<ProductRow>(stream, options =>
        {
            options.UseTheme(new ProductTheme());
            options.TemplateRows = 50;
            options.Column(x => x.Name).AllowedValues("Coffee", "Tea", "Milk");
        });

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        var sheet = workbook.Worksheet("Products");

        Assert.Equal("Id", sheet.Cell(1, 1).GetString());
        Assert.Equal("Status", sheet.Cell(1, 6).GetString());
        Assert.True(sheet.Cell(1, 1).Style.Font.Bold);
        Assert.Equal(18d, sheet.Column(4).Width);
        Assert.True(sheet.DataValidations.Any());
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

public enum ProductStatus
{
    Active,
    Inactive,
    Discontinued
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

    [ExcelColumn("Status", Order = 6)]
    public ProductStatus Status { get; set; }
}
