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
    public void Import_ShouldAcceptAliasesCaseInsensitiveOrderAndMissingOptionalColumns()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Products");
            sheet.Cell(1, 1).Value = "PRODUCT NAME";
            sheet.Cell(1, 2).Value = "id";
            sheet.Cell(2, 1).Value = "Coffee";
            sheet.Cell(2, 2).Value = 7;
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var result = new ExcelMapper().Import<ProductRow>(stream);

        Assert.Single(result);
        Assert.Equal(7, result[0].Id);
        Assert.Equal("Coffee", result[0].Name);
        Assert.Equal(0m, result[0].Price);
    }

    [Fact]
    public void Import_ColumnValidation_ShouldAssociateErrorWithMappedHeader()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Products");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(1, 3).Value = "Price";
            sheet.Cell(2, 1).Value = 1;
            sheet.Cell(2, 2).Value = "Coffee";
            sheet.Cell(2, 3).Value = -1;
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var result = new ExcelMapper().Import<ProductRow>(stream, options =>
        {
            options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
            options.Column(x => x.Price).Validate(value => value >= 0, "Price cannot be negative.");
        });

        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal("Price", error.Column);
    }

    [Fact]
    public void Import_ShouldMatchCanonicalHeaderCaseInsensitive()
    {
        using var stream = CreateProductsWorkbook(("NAME", "Coffee"), ("ID", 7));
        var result = new ExcelMapper().Import<ProductRow>(stream);

        var item = Assert.Single(result);
        Assert.Equal(7, item.Id);
        Assert.Equal("Coffee", item.Name);
    }

    [Fact]
    public void Import_ShouldMatchAliasCaseInsensitive()
    {
        using var stream = CreateProductsWorkbook(("description", "Coffee"), ("Id", 7));
        var result = new ExcelMapper().Import<ProductRow>(stream);

        Assert.Equal("Coffee", Assert.Single(result).Name);
    }

    [Fact]
    public void Import_ShouldMapColumnsRegardlessOfOrder()
    {
        using var stream = CreateProductsWorkbook(("Name", "Coffee"), ("Id", 7));
        var result = new ExcelMapper().Import<ProductRow>(stream);

        var item = Assert.Single(result);
        Assert.Equal(7, item.Id);
        Assert.Equal("Coffee", item.Name);
    }

    [Fact]
    public void Import_ShouldAllowMissingOptionalColumn()
    {
        using var stream = CreateProductsWorkbook(("Id", 7), ("Name", "Coffee"));
        var item = Assert.Single(new ExcelMapper().Import<ProductRow>(stream));

        Assert.Equal(0m, item.Price);
        Assert.False(item.Active);
    }

    [Fact]
    public void Import_RequiredColumn_ShouldAcceptAlias()
    {
        using var stream = CreateProductsWorkbook(("Id", 7), ("Product Name", "Coffee"));
        var item = Assert.Single(new ExcelMapper().Import<ProductRow>(stream));

        Assert.Equal("Coffee", item.Name);
    }

    [Fact]
    public void Import_EmptyRowBehaviorError_ShouldCollectEmptyRow()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Products");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(2, 1).Value = 1;
            sheet.Cell(2, 2).Value = "Coffee";
            sheet.Cell(3, 1).Style.Font.Bold = true;
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var result = new ExcelMapper().Import<ProductRow>(stream, options =>
        {
            options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
            options.EmptyRowBehavior = ExcelEmptyRowBehavior.Error;
        });

        Assert.Single(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(3, error.Row);
    }


    [Fact]
    public void Export_AutoFitColumns_ShouldConsiderDataAndRespectExplicitWidth()
    {
        var mapper = new ExcelMapper();
        using var stream = new MemoryStream();

        mapper.Export(new[]
        {
            new ProductRow { Id = 1, Name = "A product name much longer than its header" }
        }, stream, options =>
        {
            options.AutoFitColumns = true;
            options.Column(x => x.Id).Width(22);
        });

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        var sheet = workbook.Worksheet("Products");

        Assert.Equal(22d, sheet.Column(1).Width);
        Assert.True(sheet.Column(2).Width > "Name".Length);
    }

    [Fact]
    public void Export_ShouldApplyNativeColumnFormatsFromMappedTypes()
    {
        var mapper = new ExcelMapper();
        using var stream = new MemoryStream();

        mapper.Export(new[]
        {
            new TypedRow
            {
                When = new DateTime(2026, 9, 29, 14, 30, 0),
                Amount = 123.45m,
                Count = 7
            }
        }, stream);

        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        var sheet = workbook.Worksheet("Typed");

        Assert.Equal(XLDataType.DateTime, sheet.Cell(2, 1).DataType);
        Assert.Equal("yyyy-mm-dd hh:mm:ss", sheet.Column(1).Style.NumberFormat.Format);
        Assert.Equal(XLDataType.Number, sheet.Cell(2, 2).DataType);
        Assert.Equal("#,##0.########", sheet.Column(2).Style.NumberFormat.Format);
        Assert.Equal("0", sheet.Column(3).Style.NumberFormat.Format);
    }

    [Fact]
    public void Import_ErrorWhen_ShouldTreatTrueAsAnError()
    {
        using var stream = CreateProductsWorkbook(("Id", 7), ("Name", "Coffee"), ("Price", -1));
        var result = new ExcelMapper().Import<ProductRow>(stream, options =>
        {
            options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
            options.Column(x => x.Price).ErrorWhen(value => value < 0, "Price cannot be negative.");
        });

        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal("Price", error.Column);
        Assert.Contains("negative", error.Message);
    }

    [Fact]
    public void Import_MaxInvalidRows_ShouldStopAfterConfiguredNumberOfInvalidRows()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Products");
            sheet.Cell(1, 1).Value = "Id";
            sheet.Cell(1, 2).Value = "Name";
            sheet.Cell(1, 3).Value = "Price";

            for (var row = 2; row <= 20; row++)
            {
                sheet.Cell(row, 1).Value = row;
                sheet.Cell(row, 2).Value = $"Product {row}";
                sheet.Cell(row, 3).Value = -1;
            }

            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        var result = new ExcelMapper().Import<ProductRow>(stream, options =>
        {
            options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
            options.MaxInvalidRows = 3;
            options.Column(x => x.Price).ErrorWhen(value => value < 0, "Invalid price.");
            options.Validate(row => row.Name.Length > 0, "Name is required.");
        });

        Assert.True(result.ReachedInvalidRowLimit);
        Assert.Equal(3, result.Errors.Select(error => error.Row).Distinct().Count());
        Assert.Empty(result.Items);
    }


    [Fact]
    public void Import_Normalize_ShouldRunBeforeValidation()
    {
        using var stream = CreateProductsWorkbook(("Id", 7), ("Name", "  Coffee  "));
        var result = new ExcelMapper().Import<ProductRow>(stream, options =>
        {
            options.ErrorBehavior = ExcelImportErrorBehavior.Collect;
            options.Column(x => x.Name)
                .Normalize(value => value.Trim().PadLeft(10, '0'))
                .Validate(value => value == "0000Coffee", "Name was not normalized.");
        });

        var item = Assert.Single(result.Items);
        Assert.Equal("0000Coffee", item.Name);
        Assert.Empty(result.Errors);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Import_Normalize_ShouldSkipBlankStrings(string value)
    {
        using var stream = CreateProductsWorkbook(("Id", 7), ("Name", value));
        var calls = 0;

        var result = new ExcelMapper().Import<ProductRow>(stream, options =>
        {
            options.EmptyRowBehavior = ExcelEmptyRowBehavior.Include;
            options.Column(x => x.Name).Normalize(text =>
            {
                calls++;
                return text.Trim();
            });
        });

        Assert.Equal(0, calls);
        Assert.Equal(value, Assert.Single(result).Name);
    }

    private static MemoryStream CreateProductsWorkbook(params (string Header, object Value)[] columns)
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.AddWorksheet("Products");
            for (var i = 0; i < columns.Length; i++)
            {
                sheet.Cell(1, i + 1).Value = columns[i].Header;
                sheet.Cell(2, i + 1).Value = XLCellValue.FromObject(columns[i].Value);
            }

            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        return stream;
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

    [ExcelColumn("Name", Order = 2, Required = true, Aliases = new[] { "Product Name", "Description" })]
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


[ExcelSheet("Typed")]
public partial class TypedRow
{
    [ExcelColumn("When", Order = 1)]
    public DateTime When { get; set; }

    [ExcelColumn("Amount", Order = 2)]
    public decimal Amount { get; set; }

    [ExcelColumn("Count", Order = 3)]
    public int Count { get; set; }
}
