using System;
using System.Reflection;
using COMInteraction.InterfaceApplication.ReadValueConverters;

namespace UnitTests.InterfaceApplication.Helpers
{
    public class ActionlessReadValueConverter : IReadValueConverter
    {
        public object Convert(PropertyInfo property, object value)
        {
            return value;
        }
        public object Convert(MethodInfo method, object value)
        {
            return value;
        }
    }
}
