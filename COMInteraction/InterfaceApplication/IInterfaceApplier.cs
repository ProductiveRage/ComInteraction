using System;
namespace COMInteraction.InterfaceApplication
{
    public interface IInterfaceApplier
    {
        /// <summary>
        /// This will always be an interface, never a class
        /// </summary>
        Type TargetType { get; }

        /// <summary>
        /// This will raise an exception for null src
        /// </summary>
        object Apply(object src);
    }

    public interface IInterfaceApplier<T> : IInterfaceApplier
    {
        /// <summary>
        /// This will raise an exception for null src (T will always be an interface, never a class)
        /// </summary>
        new T Apply(object src);
    }
}
