using System;

namespace EPPlus.TableGrid.Exceptions
{
    public class RequiredPropertyException : Exception
    {
        public RequiredPropertyException(string propertyName, Type classType)
            : base($"{propertyName} is required for {classType.Name}")
        {
        }
    }
}