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
    public void ExcelEnumerable_CountsThree_AfterReading()
    {
      var excelFileStream = GetExcelFileStream();
      List<ExampleRow> itemsFromExcel;

      using (var excelEnumerable = new XcelEnumerable<ExampleRow>(excelFileStream))
      {
        itemsFromExcel = excelEnumerable.ToList();
      }

      Assert.Equal(3, itemsFromExcel.Count);
    }

    [Fact]
    public void ExcelEnumerable_MapsCorrectData_AfterReading()
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

    [Fact]
    public void ExcelEnumerable_ReadsDataFromCorrectSheet_WhenSheetNameProvided()
    {
      var excelFileStream = GetExcelFileStream();
      List<ExampleAnotherSheetRow> anotherSheetRows;

      using (var excelEnumerable = new XcelEnumerable<ExampleAnotherSheetRow>(excelFileStream, "Another Sheet"))
      {
        anotherSheetRows = excelEnumerable.ToList();
      }
      
      Assert.Equal(4, anotherSheetRows.Count);
      Assert.Equal("New York", anotherSheetRows.First().City);
    }

    [Fact]
    public void ExcelEnumerable_ThrowsInvalidOperationException_WhenSheetNameNotFound()
    {
      var excelFileStream = GetExcelFileStream();

      using (var excelEnumerable = new XcelEnumerable<ExampleAnotherSheetRow>(excelFileStream, "Unknown"))
      {
        Assert.Throws<InvalidOperationException>(excelEnumerable.ToList);
      }
    }

    private Stream GetExcelFileStream()
    {
      return Assembly.GetExecutingAssembly()
        .GetManifestResourceStream("ExcelEnumerable.Tests.Example.xlsx");
    }
  }
}