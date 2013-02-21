using System;
using System.Collections.Generic;
using COMInteraction.InterfaceApplication.ReadValueConverters;

namespace COMInteraction.InterfaceApplication
{
	/// <summary>
	/// This will cache the InterfaceApplier instances returned so that the same work need not be done multiple times. It uses basic locking and so early requests may
	/// duplicate work but later requests will not (and a lock will be required for every read in this implementation).
	/// </summary>
    public class CachingInterfaceApplierFactory : IInterfaceApplierFactory
    {
		private readonly Dictionary<Type, IInterfaceApplier> _cache;
		private readonly IInterfaceApplierFactory _interfaceApplierFactory;
		public CachingInterfaceApplierFactory(IInterfaceApplierFactory interfaceApplierFactory)
		{
			if (interfaceApplierFactory == null)
				throw new ArgumentNullException("interfaceApplierFactory");

			_cache = new Dictionary<Type, IInterfaceApplier>();
			_interfaceApplierFactory = interfaceApplierFactory;
		}

        /// <summary>
        /// Try to generate an InterfaceApplier for the specified interface (the typeparam MUST be an interface, not a class)
        /// </summary>
        /// <typeparam name="T">This will always be an interface, never a class</typeparam>
		public IInterfaceApplier<T> GenerateInterfaceApplier<T>(IReadValueConverter readValueConverter)
		{
			if (!typeof(T).IsInterface)
				throw new ArgumentException("typeparam must be an interface type", "targetInterface");
			if (readValueConverter == null)
				throw new ArgumentNullException("readValueConverter");

			lock (_cache)
			{
				if (_cache.ContainsKey(typeof(T)))
					return (IInterfaceApplier<T>)_cache[typeof(T)];
			}
			var interfaceApplier = _interfaceApplierFactory.GenerateInterfaceApplier<T>(readValueConverter);
			lock (_cache)
			{
				if (!_cache.ContainsKey(typeof(T)))
					_cache.Add(typeof(T), interfaceApplier);
			}
			return interfaceApplier;
		}

        /// <summary>
        /// Try to generate an InterfaceApplier for the specified interface (targetType MUST be an interface, not a class)
        /// </summary>
		public IInterfaceApplier GenerateInterfaceApplier(Type targetInterface, IReadValueConverter readValueConverter)
		{
			if (!targetInterface.IsInterface)
				throw new ArgumentException("must be an interface type", "targetInterface");
			if (readValueConverter == null)
				throw new ArgumentNullException("readValueConverter");

			// We'll pass all requests to this method signature to the generic one so that we know that all entries in the cache are generic IInterfaceApplier
			// and none are the non-typeparam'd IInterfaceApplier variant
			var generate = this.GetType().GetMethod("GenerateInterfaceApplier", new[] { typeof(IReadValueConverter) });
			var generateGeneric = generate.MakeGenericMethod(targetInterface);
			return (IInterfaceApplier)generateGeneric.Invoke(this, new[] { readValueConverter });
		}
	}
}
