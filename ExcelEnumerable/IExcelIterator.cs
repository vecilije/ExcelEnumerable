using System;
using System.Collections.Generic;

namespace ExcelEnumerable
{
  /// <summary>
  /// Interface which represents a reader/mapper from excel file
  /// </summary>
  /// <typeparam name="T">The type which represents </typeparam>
  public interface IExcelIterator<out T> : IEnumerable<T>, IDisposable where T : class, new()
  {
  }
}