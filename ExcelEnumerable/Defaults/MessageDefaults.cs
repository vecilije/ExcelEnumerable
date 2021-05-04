namespace ExcelEnumerable.Defaults
{
  internal static class MessageDefaults
  {
    public static string CannotMapColumnByName(string columnName) => $"Cannot map column index by name '{columnName}'";

    public static string SheetNotFound(string sheetName) => $"The sheet '{sheetName}' cannot be found.";
    
    public const string InvalidPropertyExpression = "Invalid property expression.";

    public static string PropertyNotFound<T>(string propertyName) where T : class, new() =>
      $"Property '{propertyName}' cannot be found for type '{typeof(T).Name}'";

    public static string InvalidColumnIndex(int columnIndex) => $"Configured column index '{columnIndex}' is invalid.";

    public static string InvalidColumnName(string columnName) => $"Configured column name '{columnName}' is invalid.";
    
    public const string SheetNameEmpty = "Configured sheet name is empty.";

    public const string PasswordEmpty = "Configured woorkbook password is empty.";

    public const string CannotReadColumnNames = "Unable to read column names.";
    
    public const string EmptyColumnNameFound = "Empty column name found.";
    
    public const string XcelEnumerableObsolete = "XcelEnumerable is obsolete and will be removed in future versions. " +
                                                 "Consider using ExcelIteratorCreator instead.";

    public const string XcelEnumerableColumnAttributeObsolete = "XcelEnumerableColumnAttribute is obsolete and will " +
                                                                "be removed in future versions. Property mapping is " +
                                                                "defined with configuration.";
  }
}