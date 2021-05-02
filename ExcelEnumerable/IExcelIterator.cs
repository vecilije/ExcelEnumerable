using System;
using System.Collections.Generic;

namespace ExcelEnumerable
{
  public interface IExcelIterator<T> : IEnumerable<T>, IDisposable where T : class, new()
  {
  }
}