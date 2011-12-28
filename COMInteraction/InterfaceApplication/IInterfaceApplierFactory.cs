using System;
using COMInteraction.InterfaceApplication.ReadValueConverters;

namespace COMInteraction.InterfaceApplication
{
    public interface IInterfaceApplierFactory
    {
        /// <summary>
        /// Try to generate an InterfaceApplier for the specified interface (the typeparam MUST be an interface, not a class)
        /// </summary>
        /// <typeparam name="T">This will always be an interface, never a class</typeparam>
        IInterfaceApplier<T> GenerateInterfaceApplier<T>(IReadValueConverter readValueConverter);

        /// <summary>
        /// Try to generate an InterfaceApplier for the specified interface (targetType MUST be an interface, not a class)
        /// </summary>
        IInterfaceApplier GenerateInterfaceApplier(Type targetType, IReadValueConverter readValueConverter);
    }
}
