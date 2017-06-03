using System;

namespace EPPlus.TableGrid.Core.Exceptions
{
    public class RequiredPropertyException : Exception
    {
        public RequiredPropertyException(string propertyName, Type classType)
            : base($"{propertyName} is required for {classType.Name}")
        {
        }
    }
}