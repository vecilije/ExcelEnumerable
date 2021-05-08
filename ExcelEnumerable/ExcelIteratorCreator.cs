using System;
using System.IO;
using ExcelEnumerable.Configuration;

namespace ExcelEnumerable
{
  /// <summary>
  /// Creator of <see cref="IExcelIterator{T}"/> instance
  /// </summary>
  public static class ExcelIteratorCreator
  {
    /// <summary>
    /// Creates a new instance of <see cref="IExcelIterator{T}"/> with optional action for defining configuration
    /// </summary>
    /// <param name="fileStream">The stream of an excel file</param>
    /// <param name="builderAction">The optional action for defining configuration
    /// If not specified, the default configuration will be used.</param>
    /// <typeparam name="T">The type which represents an item (row) from excel file</typeparam>
    /// <returns>A new instance of <see cref="IExcelIterator{T}"/></returns>
    public static IExcelIterator<T> Create<T>(Stream fileStream, 
      Action<ExcelIteratorConfigurationBuilder<T>> builderAction = null) 
      where T : class, new()
    {
      var configurationBuilder = ExcelIteratorConfigurationBuilder<T>.Default();
      
      builderAction?.Invoke(configurationBuilder);

      return new ExcelIterator<T>(fileStream, configurationBuilder.Build());
    }

    /// <summary>
    /// Creates a new instance of <see cref="IExcelIterator{T}"/> with optional configuration
    /// </summary>
    /// <param name="fileStream">The stream of an excel file</param>
    /// <param name="configuration">The optional configuration.
    /// If not specified the default configuration will be used</param>
    /// <typeparam name="T">The type which represents an item (row) from excel file</typeparam>
    /// <returns>A new instance of <see cref="IExcelIterator{T}"/></returns>
    public static IExcelIterator<T> Create<T>(Stream fileStream, ExcelIteratorConfiguration configuration = null)
      where T : class, new()
    {
      return new ExcelIterator<T>(fileStream, configuration ?? ExcelIteratorConfigurationBuilder<T>.Default().Build());
    }
      
  }
}