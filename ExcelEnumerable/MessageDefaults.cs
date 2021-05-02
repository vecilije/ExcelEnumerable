namespace ExcelEnumerable
{
  internal static class MessageDefaults
  {
    public static string CannotMapColumnByName(string columnName) => $"Cannot map column index by name '{columnName}'";

    public static string SheetNameNotFound(string sheetName) => $"The sheet '{sheetName}' cannot be found.";
    public const string InvalidPropertyExpression = "Invalid property expression.";

    public static string PropertyNotFound<T>(string propertyName) where T : class, new() =>
      $"Property '{propertyName}' cannot be found for type '{typeof(T).Name}'";

    public static string InvalidColumnIndex(int columnIndex) => $"Configured column index '{columnIndex}' is invalid.";

    public static string InvalidColumnName(string columnName) => $"Configured column name '{columnName}' is invalid.";
    public const string InvalidConvertValueFunction = "Configured source value conversion function is not valid.";
  }
}