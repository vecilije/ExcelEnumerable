using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace ExcelEnumerable.Tests
{
  public class ExcelIteratorTests
  {
    public ExcelIteratorTests()
    {
      System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }

    [Fact]
    public void Iterator_MapsCorrectNumberOfItems_WhenMixedConfiguration()
    {
      using var iterator = ExcelIteratorCreator.Create<ExampleRow>(GetFileStream("Example.xlsx"),
        builder =>
        {
          builder.FirstRowContainsColumnNames(flag: true);
          builder.SkipEmptyColumnNames(flag: true);
          builder.MapByName(p => p.FirstName, "First NAME");
          builder.MapByIndex(p => p.LastName, 1);
          builder.Ignore(p => p.Address);
          builder.MapByName(p => p.IsActive, "is AcTIve");
          builder.ConvertSourceValue(p => p.IsActive, sourceValue => sourceValue?.ToString() == "t");
        });

      var items = iterator.ToList();

      Assert.Equal(3, items.Count);
    }

    [Fact]
    public void Iterator_MapsCorrectData_WhenMixedConfiguration()
    {
      using var iterator = ExcelIteratorCreator.Create<ExampleRow>(GetFileStream("Example.xlsx"),
        builder =>
        {
          builder.FirstRowContainsColumnNames(flag: true);
          builder.SkipEmptyColumnNames(flag: true);
          builder.MapByName(p => p.FirstName, "First NAME");
          builder.MapByIndex(p => p.LastName, 2);
          builder.Ignore(p => p.Address);
          builder.MapByName(p => p.IsActive, "is AcTIve");
          builder.ConvertSourceValue(p => p.IsActive, sourceValue => sourceValue?.ToString() == "t");
        });

      var firstItem = iterator.First();

      Assert.Equal(1, firstItem.Id);
      Assert.Equal("Unknown", firstItem.FirstName);
      Assert.Equal("User", firstItem.LastName);
      Assert.True(string.IsNullOrWhiteSpace(firstItem.Address));
      Assert.Equal(new DateTime(1990, 3, 7), firstItem.Date);
      Assert.True(firstItem.IsActive);
    }

    [Fact]
    public void Iterator_ReadsCorrectData_WhenMappedByName()
    {
      using var iterator = ExcelIteratorCreator.Create<ExampleRow>(GetFileStream("Example.xlsx"),
        builder =>
        {
          builder.FirstRowContainsColumnNames(flag: true);

          builder.MapByName(p => p.FirstName, "First Name");
          builder.MapByName(p => p.LastName, "last name");
          builder.MapByName(p => p.IsActive, "Is Active");
          builder.ConvertSourceValue(p => p.IsActive, sourceValue => sourceValue?.ToString() == "t");
        });

      var firstItem = iterator.First();

      Assert.Equal(1, firstItem.Id);
      Assert.Equal("Unknown", firstItem.FirstName);
      Assert.Equal("User", firstItem.LastName);
      Assert.Equal("Sample Address 11", firstItem.Address);
      Assert.Equal(new DateTime(1990, 3, 7), firstItem.Date);
      Assert.True(firstItem.IsActive);
    }

    [Fact]
    public void Iterator_ReadsCorrectData_WhenMappedByIndex()
    {
      using var iterator = ExcelIteratorCreator.Create<ExampleRow>(GetFileStream("Example.xlsx"),
        builder =>
        {
          builder.FirstRowContainsColumnNames(flag: true);

          builder.MapByIndex(p => p.Id, 0);
          builder.MapByIndex(p => p.FirstName, 1);
          builder.MapByIndex(p => p.LastName, 2);
          builder.MapByIndex(p => p.Address, 3);
          builder.MapByIndex(p => p.Date, 4);
          builder.MapByIndex(p => p.IsActive, 5);
          builder.ConvertSourceValue(p => p.IsActive, sourceValue => sourceValue?.ToString() == "t");
        });

      var firstItem = iterator.First();

      Assert.Equal(1, firstItem.Id);
      Assert.Equal("Unknown", firstItem.FirstName);
      Assert.Equal("User", firstItem.LastName);
      Assert.Equal("Sample Address 11", firstItem.Address);
      Assert.Equal(new DateTime(1990, 3, 7), firstItem.Date);
      Assert.True(firstItem.IsActive);
    }

    [Fact]
    public void Iterator_ReadsCorrectDataFromCorrectSheet_WhenSheetSpecified()
    {
      using var iterator = ExcelIteratorCreator.Create<ExampleAnotherSheetRow>(GetFileStream("Example.xlsx"),
        builder =>
        {
          builder.FirstRowContainsColumnNames(flag: true);
          builder.UseSheetName("Another Sheet");
        });

      var lastItem = iterator.Last();

      Assert.Equal("Las Vegas", lastItem.City);
    }

    [Fact]
    public void Iterator_ThrowsException_WhenSheetNotFound()
    {
      using var iterator = ExcelIteratorCreator.Create<ExampleAnotherSheetRow>(GetFileStream("Example.xlsx"),
        builder =>
        {
          builder.FirstRowContainsColumnNames(flag: true);
          builder.UseSheetName("Non-Existing Sheet");
        });

      var exception = Assert.Throws<InvalidOperationException>(() => iterator.ToList());
      Assert.Equal($"The sheet 'Non-Existing Sheet' cannot be found.", exception.Message);
    }

    [Fact]
    public void Iterator_ThrowsException_WhenNotAbleToMapColumn()
    {
      using var iterator = ExcelIteratorCreator.Create<ExampleRow>(GetFileStream("Example.xlsx"),
        builder =>
        {
          builder.FirstRowContainsColumnNames(flag: true);

          builder.MapByName(p => p.FirstName, "First Name");
          builder.MapByName(p => p.LastName, "Non-Existing Column");
          builder.MapByName(p => p.IsActive, "Is Active");
          builder.ConvertSourceValue(p => p.IsActive, sourceValue => sourceValue?.ToString() == "t");
        });

      var exception = Assert.ThrowsAny<InvalidOperationException>(() => iterator.ToList());
      Assert.Equal($"Cannot map column index by name 'Non-Existing Column'", exception.Message);
    }

    [Fact]
    public void Iterator_MapsCorrectData_WhenSkippingWhitespaceForColumnName()
    {
      using var iterator = ExcelIteratorCreator.Create<ExampleRow>(GetFileStream("Example.xlsx"),
        builder =>
        {
          builder.FirstRowContainsColumnNames(flag: true);
          builder.SkipWhitespaceForColumnNames(flag: true);
          
          builder.ConvertSourceValue(p => p.IsActive, sourceValue => sourceValue?.ToString() == "t");
        });

      var firstItem = iterator.First();

      Assert.Equal(1, firstItem.Id);
      Assert.Equal("Unknown", firstItem.FirstName);
      Assert.Equal("User", firstItem.LastName);
      Assert.Equal("Sample Address 11", firstItem.Address);
      Assert.Equal(new DateTime(1990, 3, 7), firstItem.Date);
      Assert.True(firstItem.IsActive);
    }

    private Stream GetFileStream(string fileName)
    {
      return Assembly.GetExecutingAssembly()
        .GetManifestResourceStream($"ExcelEnumerable.Tests.{fileName}");
    }
  }
}