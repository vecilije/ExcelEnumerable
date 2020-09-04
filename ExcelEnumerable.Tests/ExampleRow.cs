using System;

namespace ExcelEnumerable.Tests
{
  public class ExampleRow
  {
    [XcelEnumerableColumn("First Name")]
    public string FirstName { get; set; }
    
    [XcelEnumerableColumn("Last Name")]
    public string LastName { get; set; }

    public string Address { get; set; }
    
    public DateTime Date { get; set; }
    
    [XcelEnumerableColumn("Is Active")]
    public bool IsActive { get; set; }

    public override int GetHashCode()
    {
      return HashCode.Combine(FirstName, LastName, Address, Date, IsActive);
    }

    public override bool Equals(object obj)
    {
      if (obj == null)
        return false;
      
      if (ReferenceEquals(this, obj))
        return true;
      
      if (obj is ExampleRow objToCompare)
        return GetHashCode() == objToCompare.GetHashCode();

      return false;
    }
  }
}