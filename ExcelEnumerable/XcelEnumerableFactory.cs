using System.IO;

namespace ExcelEnumerable
{
  public class XcelEnumerableFactory : IExcelEnumerableFactory
  {
    public IExcelEnumerable<TRow> Create<TRow>(string fileName) where TRow : class
    {
      return new XcelEnumerable<TRow>(fileName);
    }

    public IExcelEnumerable<TRow> Create<TRow>(Stream fileStream) where TRow : class
    {
      return new XcelEnumerable<TRow>(fileStream);
    }
  }
}