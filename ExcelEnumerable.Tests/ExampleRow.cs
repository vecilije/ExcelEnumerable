using System;

namespace ExcelEnumerable.Tests
{
  public class ExampleRow
  {
    public int Id { get; set; }
    
    public string FirstName { get; set; }
    
    public string LastName { get; set; }

    public string Address { get; set; }
    
    public DateTime Date { get; set; }
    
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