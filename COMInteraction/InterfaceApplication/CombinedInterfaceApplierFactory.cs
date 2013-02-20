using System;
using COMInteraction.InterfaceApplication.ReadValueConverters;
using COMInteraction.Misc;

namespace COMInteraction.InterfaceApplication
{
	/// <summary>
	/// This will allow for interface application by three different manners; using Reflection, through the IDispatch interface or straight-through if a source
	/// reference already implements the required interface. More than interface applier may be used in a single IInterfaceApplier.Apply call if the source
	/// object has properties or method return types that require additional interfaces applying to them (some may require IDispatch-wrapping, some may
	/// work through Reflection, some may already be the required type). If a conversion is encountered that none of these methods can deal with then
	/// an exception will be raised.
	/// </summary>
	public class CombinedInterfaceApplierFactory : IInterfaceApplierFactory
	{
		private readonly IInterfaceApplierFactory _reflectionFactory;
		private readonly IInterfaceApplierFactory _idispatchFactory;
		public CombinedInterfaceApplierFactory(IInterfaceApplierFactory reflectionFactory, IInterfaceApplierFactory idispatchFactory)
		{
			if (reflectionFactory == null)
				throw new ArgumentNullException("reflectionFactory");
			if (idispatchFactory == null)
				throw new ArgumentNullException("idispatchFactory");

			_reflectionFactory = reflectionFactory;
			_idispatchFactory = idispatchFactory;
		}

		/// <summary>
		/// Try to generate an InterfaceApplier for the specified interface (the typeparam MUST be an interface, not a class)
		/// </summary>
		/// <typeparam name="T">This will always be an interface, never a class</typeparam>
		public IInterfaceApplier<T> GenerateInterfaceApplier<T>(IReadValueConverter readValueConverter)
		{
			if (!typeof(T).IsInterface)
				throw new ArgumentException("typeparam must be an interface type");
			if (readValueConverter == null)
				throw new ArgumentNullException("readValueConverter");

			return new InterfaceApplier<T>(
				new DelayedExecutor<IInterfaceApplier<T>>(
					() => _reflectionFactory.GenerateInterfaceApplier<T>(readValueConverter)
				),
				new DelayedExecutor<IInterfaceApplier<T>>(
					() => _idispatchFactory.GenerateInterfaceApplier<T>(readValueConverter)
				)
			);
		}

		/// <summary>
		/// Try to generate an InterfaceApplier for the specified interface (targetType MUST be an interface, not a class)
		/// </summary>
		public IInterfaceApplier GenerateInterfaceApplier(Type targetInterface, IReadValueConverter readValueConverter)
		{
			if (targetInterface == null)
				throw new ArgumentNullException("targetInterface");
			if (!targetInterface.IsInterface)
				throw new ArgumentException("typeparam must be an interface type", "targetInterface");
			if (readValueConverter == null)
				throw new ArgumentNullException("readValueConverter");

			return new InterfaceApplier(
				targetInterface,
				new DelayedExecutor<IInterfaceApplier>(
					() => _reflectionFactory.GenerateInterfaceApplier(targetInterface, readValueConverter)
				),
				new DelayedExecutor<IInterfaceApplier>(
					() => _idispatchFactory.GenerateInterfaceApplier(targetInterface, readValueConverter)
				)
			);
		}

		private class InterfaceApplier<T> : IInterfaceApplier<T>
		{
			private readonly DelayedExecutor<IInterfaceApplier<T>> _reflectionApplier;
			private readonly DelayedExecutor<IInterfaceApplier<T>> _idispatchApplier;
			public InterfaceApplier(DelayedExecutor<IInterfaceApplier<T>> reflectionApplier, DelayedExecutor<IInterfaceApplier<T>> idispatchApplier)
			{
				if (!typeof(T).IsInterface)
					throw new ArgumentException("typeparam must be an interface type", "targetInterface");
				if (reflectionApplier == null)
					throw new ArgumentNullException("reflectionApplier");
				if (idispatchApplier == null)
					throw new ArgumentNullException("idispatchApplier");

				_reflectionApplier = reflectionApplier;
				_idispatchApplier = idispatchApplier;
			}

			/// <summary>
			/// This will raise an exception for null src
			/// </summary>
			public T Apply(object src)
			{
				if (src == null)
					throw new ArgumentNullException("src");

				if (src is T)
					return (T)src;

				return src.GetType().IsCOMObject ? _idispatchApplier.Value.Apply(src) : _reflectionApplier.Value.Apply(src);
			}

			/// <summary>
			/// This will always be an interface, never a class
			/// </summary>
			public Type TargetType { get { return typeof(T); } }

			object IInterfaceApplier.Apply(object src)
			{
				if (src == null)
					throw new ArgumentNullException("src");

				return Apply(src);
			}
		}

		private class InterfaceApplier : IInterfaceApplier
		{
			private readonly DelayedExecutor<IInterfaceApplier> _reflectionApplier;
			private readonly DelayedExecutor<IInterfaceApplier> _idispatchApplier;
			public InterfaceApplier(Type targetInterface, DelayedExecutor<IInterfaceApplier> reflectionApplier, DelayedExecutor<IInterfaceApplier> idispatchApplier)
			{
				if (targetInterface == null)
					throw new ArgumentNullException("targetInterface");
				if (!targetInterface.IsInterface)
					throw new ArgumentException("targetInterface must be an interface type", "targetInterface");
				if (reflectionApplier == null)
					throw new ArgumentNullException("reflectionApplier");
				if (idispatchApplier == null)
					throw new ArgumentNullException("idispatchApplier");

				TargetType = targetInterface;
				_reflectionApplier = reflectionApplier;
				_idispatchApplier = idispatchApplier;
			}

			/// <summary>
			/// This will raise an exception for null src
			/// </summary>
			public object Apply(object src)
			{
				if (src == null)
					throw new ArgumentNullException("src");

				var srcType = src.GetType();
				if (TargetType.IsAssignableFrom(srcType))
					return src;

				return srcType.IsCOMObject ? _idispatchApplier.Value.Apply(src) : _reflectionApplier.Value.Apply(src);
			}

			/// <summary>
			/// This will always be an interface, never a class
			/// </summary>
			public Type TargetType { get; private set; }
		}
	}
}
