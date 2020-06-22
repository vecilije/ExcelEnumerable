using System.IO;

namespace ExcelEnumerable
{
  public interface IExcelEnumerableFactory
  {
    IExcelEnumerable<TRow> Create<TRow>(string fileName) where TRow : class;
    IExcelEnumerable<TRow> Create<TRow>(Stream fileStream) where TRow : class;
  }
}