using System;
using System.Collections.Generic;

namespace ExcelEnumerable
{
  public interface IExcelEnumerable<TResult> : IEnumerable<TResult>, IDisposable
    where TResult : class
  {
  }
}