using System;
using System.Collections.Generic;
using ExcelEnumerable.Defaults;

namespace ExcelEnumerable
{
  [Obsolete(MessageDefaults.XcelEnumerableObsolete)]
  public interface IExcelEnumerable<TResult> : IEnumerable<TResult>, IDisposable
    where TResult : class
  {
  }
}