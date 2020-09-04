using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace ExcelEnumerable.Tests
{
  public class XcelEnumerableTests
  {
    public XcelEnumerableTests()
    {
      System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }

    [Fact]
    public void GetItems_ShouldCountThree()
    {
      //  Arrange
      var excelFileStream = GetExcelFileStream();
      List<ExampleRow> itemsFromExcel;

      //  Act
      using (var excelEnumerable = new XcelEnumerable<ExampleRow>(excelFileStream))
        itemsFromExcel = excelEnumerable.ToList();

      //  Assert
      Assert.Equal(3, itemsFromExcel.Count);
    }

    [Fact]
    public void GetItems_ShouldMapCorrectData()
    {
      var excelFileStream = GetExcelFileStream();
      ExampleRow itemFromExcel;

      using (var excelEnumerable = new XcelEnumerable<ExampleRow>(excelFileStream))
      {
        itemFromExcel = excelEnumerable.First();
      }

      var expectedItem = new ExampleRow
      {
        FirstName = "Unknown",
        LastName = "User",
        Address = "Sample Address 11",
        Date = new DateTime(1990, 3, 7),
        IsActive = true
      };
      
      Assert.Equal(expectedItem, itemFromExcel);
    }

    private Stream GetExcelFileStream()
    {
      return Assembly.GetExecutingAssembly()
        .GetManifestResourceStream("ExcelEnumerable.Tests.Example.xlsx");
    }
  }
}