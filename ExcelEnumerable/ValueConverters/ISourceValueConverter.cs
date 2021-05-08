namespace ExcelEnumerable.ValueConverters
{
  /// <summary>
  /// Interface for defining conversion from source value (excel) to destination value (property)
  /// </summary>
  public interface ISourceValueConverter
  {
    /// <summary>
    /// Converts value from source (excel) to destination (property)
    /// </summary>
    /// <param name="sourceValue">The value from excel file</param>
    /// <returns>The value to be set to a property</returns>
    object ConvertValue(object sourceValue);
  }
}