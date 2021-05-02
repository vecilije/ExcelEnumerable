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
    public void Iterator_ReturnsThreeRows_WhenReadingDone()
    {
      using var iterator = ExcelIteratorCreator.Create<ExampleRow>(GetExcelFileStream(), builder =>
      {
        builder.MapByName(p => p.FirstName, "First Name");
        builder.MapByName(p => p.LastName, "Last Name");
        
        builder.Ignore(p => p.Address);
        
        builder.MapByName(p => p.IsActive, "Is Active");
        builder.ConvertValue(p => p.IsActive, 
          (source) => source?.ToString() == "t");
      });

      var rows = iterator.ToList();
      
      Assert.Equal(3, rows.Count);
    }
    
    private Stream GetExcelFileStream()
    {
      return Assembly.GetExecutingAssembly()
        .GetManifestResourceStream("ExcelEnumerable.Tests.Example.xlsx");
    }
  }
}