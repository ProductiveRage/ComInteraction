﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Threading;
using COMInteraction.InterfaceApplication.ReadValueConverters;
using COMInteraction.Misc;

namespace COMInteraction.InterfaceApplication
{
	/// <summary>
	/// Return a delegate that will take an object reference and return an implementation of a specified interface - any calls to properties or methods of that interface
	/// are passed straight through to the wrapped reference using Reflection. If any of the interface's properties or methods are not available on the wrapped reference
	/// then an exception will be thrown when requested. Any interfaces that the target interface implements will be accounted for. This can be used to wrap POCO objects
	/// or COM objects, for example. Note: Only handling interfaces, not classes, means there are less things to consider - there are no protected set methods or fields,
	/// for example. This implementation wraps references that implement IDispatch and uses that interface to interact with the source data. The IInterfaceApplier
	/// implementations generated by this class will throw an exception if the Apply method is called with a reference that does not implement IDispatch.
	/// </summary>
	public class DispatchInterfaceApplierFactory : IInterfaceApplierFactory
	{
		// ================================================================================================================================
		// CLASS INITIALISATION
		// ================================================================================================================================
		private bool _createComVisibleClasses;
		private DelayedExecutor<ModuleBuilder> _moduleBuilder;
		public DispatchInterfaceApplierFactory(string assemblyName, ComVisibilityOptions comVisibilityOfClasses)
		{
			assemblyName = (assemblyName ?? "").Trim();
			if (assemblyName == "")
				throw new ArgumentException("Null or empty assemblyName specified");
			if (!Enum.IsDefined(typeof(ComVisibilityOptions), comVisibilityOfClasses))
				throw new ArgumentOutOfRangeException("comVisibilityOfClasses");

			_createComVisibleClasses = (comVisibilityOfClasses == ComVisibilityOptions.Visible);
			_moduleBuilder = new DelayedExecutor<ModuleBuilder>(
				() =>
				{
					var assemblyBuilder = Thread.GetDomain().DefineDynamicAssembly(
						new AssemblyName(assemblyName),
						AssemblyBuilderAccess.Run
					);
					return assemblyBuilder.DefineDynamicModule(
						assemblyBuilder.GetName().Name,
						false
					);
				}
			);
		}

		// ================================================================================================================================
		// IInterfaceApplierFactory IMPLEMENTATION
		// ================================================================================================================================
		/// <summary>
		/// Try to generate an InterfaceApplier for the specified interface (the targetType MUST be an interface, not a class), values returned from properties
		/// and methods will be passed through the specified readValueConverter in case manipulation is required (eg. applying an interface to particular
		/// returned property values of the targetType)
		/// </summary>
		public IInterfaceApplier GenerateInterfaceApplier(Type targetType, IReadValueConverter readValueConverter)
		{
			var generate = this.GetType().GetMethod("GenerateInterfaceApplier", new[] { typeof(IReadValueConverter) });
			var generateGeneric = generate.MakeGenericMethod(targetType);
			return (IInterfaceApplier)generateGeneric.Invoke(this, new[] { readValueConverter });
		}

		/// <summary>
		/// Try to generate an InterfaceApplier for the specified interface, values returned from properties and methods will be passed through the specified
		/// readValueConverter in case manipulation is required (eg. applying an interface to particular returned property values of the targetType)
		/// </summary>
		/// <typeparam name="T">This will always be an interface, never a class</typeparam>
		public IInterfaceApplier<T> GenerateInterfaceApplier<T>(IReadValueConverter readValueConverter)
		{
			if (!typeof(T).IsInterface)
				throw new ArgumentException("typeparam must be an interface type", "targetInterface");
			if (readValueConverter == null)
				throw new ArgumentNullException("readValueConverter");

			var typeName = "IDispatchInterfaceApplier" + Guid.NewGuid().ToString("N");
			var typeBuilder = _moduleBuilder.Value.DefineType(
				typeName,
				TypeAttributes.Public | TypeAttributes.Class | TypeAttributes.AutoClass | TypeAttributes.AnsiClass | TypeAttributes.BeforeFieldInit | TypeAttributes.AutoLayout,
				typeof(object),
				new Type[] { typeof(T) }
			);

			// ================================================================================================
			// Add [ComVisible(true)] and [ClassInterface(ClassInterfaceType.None)] attributes to the type so
			// that we can pass the generated objects through to COM object (eg. WSC controls) if desired
			// ================================================================================================
			if (_createComVisibleClasses)
			{
				typeBuilder.SetCustomAttribute(
					new CustomAttributeBuilder(
						typeof(ComVisibleAttribute).GetConstructor(new[] { typeof(bool) }),
						new object[] { true }
					)
				);
				typeBuilder.SetCustomAttribute(
					new CustomAttributeBuilder(
						typeof(ClassInterfaceAttribute).GetConstructor(new[] { typeof(ClassInterfaceType) }),
						new object[] { ClassInterfaceType.None }
					)
				);
			}

			// ================================================================================================
			// Ensure we account for any interfaces that the target interface implements (recursively)
			// ================================================================================================
			var interfaces = (new InterfaceHierarchyCombiner(typeof(T))).Interfaces;

			// ================================================================================================
			// Declare private "src" and "readValueConverter" fields
			// ================================================================================================
			var srcField = typeBuilder.DefineField("_src", typeof(object), FieldAttributes.Private);
			var readValueConverterField = typeBuilder.DefineField("_readValueConverter", typeof(IReadValueConverter), FieldAttributes.Private);

			// ================================================================================================
			// Generate the constructor, properties and methods (fields can't be declared on interfaces)
			// ================================================================================================
			generateConstructor(typeBuilder, srcField, readValueConverterField);
			generateProperties<T>(typeBuilder, srcField, readValueConverterField, interfaces);
			generateMethods<T>(typeBuilder, srcField, readValueConverterField, interfaces);

			// ================================================================================================
			// Return a new InterfaceApplier references
			// ================================================================================================
			return new InterfaceApplier<T>(
				src => (T)Activator.CreateInstance(
					typeBuilder.CreateType(),
					src,
					readValueConverter
				)
			);
		}

		private static void generateConstructor(TypeBuilder typeBuilder, FieldBuilder srcField, FieldBuilder readValueConverterField)
		{
			if (typeBuilder == null)
				throw new ArgumentNullException("typeBuilder");
			if (srcField == null)
				throw new ArgumentNullException("srcField");
			if (readValueConverterField == null)
				throw new ArgumentNullException("readValueConverterField");
			if (!typeof(IReadValueConverter).IsAssignableFrom(readValueConverterField.FieldType))
				throw new ArgumentException("readValueConverterField must be assignable to IReadValueConverter");

			var ctorBuilder = typeBuilder.DefineConstructor(
				MethodAttributes.Public,
				CallingConventions.Standard,
				new[]
                {
                    typeof(object),
                    typeof(IReadValueConverter)
                }
			);
			var ilCtor = ctorBuilder.GetILGenerator();

			// Generate: base.ctor()
			ilCtor.Emit(OpCodes.Ldarg_0);
			ilCtor.Emit(OpCodes.Call, typeBuilder.BaseType.GetConstructor(Type.EmptyTypes));

			// Generate: if (src != null), don't throw new ArgumentException("src")
			var nonNullSrcArgumentLabel = ilCtor.DefineLabel();
			ilCtor.Emit(OpCodes.Ldarg_1);
			ilCtor.Emit(OpCodes.Brtrue, nonNullSrcArgumentLabel);
			ilCtor.Emit(OpCodes.Ldstr, "src");
			ilCtor.Emit(OpCodes.Newobj, typeof(ArgumentNullException).GetConstructor(new[] { typeof(string) }));
			ilCtor.Emit(OpCodes.Throw);
			ilCtor.MarkLabel(nonNullSrcArgumentLabel);

			// Generate: if (IDispatchAccess.ImplementsIDispatch(src)), don't throw new ArgumentException("src does not implement IDispatch")
			var implementsIDispatchMethodInformation = typeof(IDispatchAccess).GetMethod("ImplementsIDispatch");
			if (implementsIDispatchMethodInformation == null)
				throw new Exception("Failed to locate IDispatchAccess.ImplementsIDispatch");
			var validIDispatchReferenceLabel = ilCtor.DefineLabel();
			ilCtor.Emit(OpCodes.Ldarg_1);
			ilCtor.Emit(OpCodes.Call, implementsIDispatchMethodInformation);
			ilCtor.Emit(OpCodes.Brtrue, validIDispatchReferenceLabel);
			ilCtor.Emit(OpCodes.Ldstr, "src does not implement IDispatch");
			ilCtor.Emit(OpCodes.Newobj, typeof(ArgumentException).GetConstructor(new[] { typeof(string) }));
			ilCtor.Emit(OpCodes.Throw);
			ilCtor.MarkLabel(validIDispatchReferenceLabel);

			// Generate: if (readValueConverter != null), don't throw new ArgumentException("readValueConverter")
			var nonNullReadValueConverterArgumentLabel = ilCtor.DefineLabel();
			ilCtor.Emit(OpCodes.Ldarg_2);
			ilCtor.Emit(OpCodes.Brtrue, nonNullReadValueConverterArgumentLabel);
			ilCtor.Emit(OpCodes.Ldstr, "readValueConverter");
			ilCtor.Emit(OpCodes.Newobj, typeof(ArgumentNullException).GetConstructor(new[] { typeof(string) }));
			ilCtor.Emit(OpCodes.Throw);
			ilCtor.MarkLabel(nonNullReadValueConverterArgumentLabel);

			// Generate: this._src = src
			ilCtor.Emit(OpCodes.Ldarg_0);
			ilCtor.Emit(OpCodes.Ldarg_1);
			ilCtor.Emit(OpCodes.Stfld, srcField);

			// Generate: this._readValueConverter = readValueConverter
			ilCtor.Emit(OpCodes.Ldarg_0);
			ilCtor.Emit(OpCodes.Ldarg_2);
			ilCtor.Emit(OpCodes.Stfld, readValueConverterField);

			// All done
			ilCtor.Emit(OpCodes.Ret);
		}

		private static void generateProperties<T>(TypeBuilder typeBuilder, FieldBuilder srcField, FieldBuilder readValueConverterField, NonNullImmutableList<Type> interfaces)
		{
			if (!typeof(T).IsInterface)
				throw new ArgumentException("typeparam must be an interface type", "targetInterface");
			if (typeBuilder == null)
				throw new ArgumentNullException("typeBuilder");
			if (srcField == null)
				throw new ArgumentNullException("srcField");
			if (readValueConverterField == null)
				throw new ArgumentNullException("readValueConverterField");
			if (!typeof(IReadValueConverter).IsAssignableFrom(readValueConverterField.FieldType))
				throw new ArgumentException("readValueConverterField must be assignable to IReadValueConverter");
			if (interfaces == null)
				throw new ArgumentNullException("interfaces");

			// Since we're only dealing with interfaces, we need only conside the CanRead and CanWrite values on the properties, nothing more complicated)
			// - There could be multiple properties defined with different types if they appear in multiple interfaces, this doesn't actually cause any
			//   problems so we won't bother doing any work to prevent it happening
			foreach (var property in interfaces.SelectMany(i => i.GetProperties()))
			{
				// Prepare the property we'll add get and/or set accessors to
				var propBuilder = typeBuilder.DefineProperty(
					property.Name,
					PropertyAttributes.None,
					property.PropertyType,
					Type.EmptyTypes
				);

				// Define get method, if required
				if (property.CanRead)
				{
					var getFuncBuilder = typeBuilder.DefineMethod(
						"get_" + property.Name,
						MethodAttributes.Public | MethodAttributes.HideBySig | MethodAttributes.NewSlot | MethodAttributes.SpecialName | MethodAttributes.Virtual | MethodAttributes.Final,
						property.PropertyType,
						Type.EmptyTypes
					);
					WritePropertyGetter(getFuncBuilder, srcField, readValueConverterField, property);
					propBuilder.SetGetMethod(getFuncBuilder);
				}

				// Define set method, if required
				if (property.CanWrite)
				{
					var setFuncBuilder = typeBuilder.DefineMethod(
						"set_" + property.Name,
						MethodAttributes.Public | MethodAttributes.HideBySig | MethodAttributes.SpecialName | MethodAttributes.Virtual,
						null,
						new Type[] { property.PropertyType }
					);
					WritePropertySetter(setFuncBuilder, srcField, property);
					propBuilder.SetSetMethod(setFuncBuilder);
				}
			}
		}

		private static void WritePropertyGetter(MethodBuilder funcBuilder, FieldBuilder srcField, FieldBuilder readValueConverterField, PropertyInfo property)
		{
			if (funcBuilder == null)
				throw new ArgumentNullException("funcBuilder");
			if (srcField == null)
				throw new ArgumentNullException("srcField");
			if (readValueConverterField == null)
				throw new ArgumentNullException("readValueConverterField");
			if (property == null)
				throw new ArgumentNullException("property");

			// Generate: return this._readValueConverter.Convert(
			//  property.DeclaringType.GetProperty(property.Name)
			//  IDispatchAccess.GetProperty(this._src, method.Name, indices)
			// );
			var ilGetFunc = funcBuilder.GetILGenerator();

			// Generate: object[] indexParameters
			var indexParameterValues = ilGetFunc.DeclareLocal(typeof(object[]));

			// Generate: indexParameters = new object[x]
			var indexParameters = property.GetIndexParameters();
			var indexParameterTypes = new List<Type>();
			foreach (var indexParameter in indexParameters)
			{
				// These options aren't supported by C# but to be on the safe side we'll ensure they're not enabled some other way
				if (indexParameter.IsOut)
					throw new ArgumentException("Output parameters are not supported");
				if (indexParameter.IsOptional)
					throw new ArgumentException("Optional parameters are not supported");
				if (indexParameter.ParameterType.IsByRef)
					throw new ArgumentException("Ref parameters are not supported");
				indexParameterTypes.Add(indexParameter.ParameterType);
			}
			ilGetFunc.Emit(OpCodes.Ldc_I4, indexParameters.Length);
			ilGetFunc.Emit(OpCodes.Newarr, typeof(Object));
			ilGetFunc.Emit(OpCodes.Stloc_0);
			for (var index = 0; index < indexParameters.Length; index++)
			{
				// Generate: indexParameters[n] = ..;
				var parameter = indexParameters[index];
				ilGetFunc.Emit(OpCodes.Ldloc_0);
				ilGetFunc.Emit(OpCodes.Ldc_I4, index);
				ilGetFunc.Emit(OpCodes.Ldarg, index + 1);
				if (parameter.ParameterType.IsValueType)
					ilGetFunc.Emit(OpCodes.Box, parameter.ParameterType);
				ilGetFunc.Emit(OpCodes.Stelem_Ref);
			}

			ilGetFunc.Emit(OpCodes.Ldarg_0);
			ilGetFunc.Emit(OpCodes.Ldfld, readValueConverterField);
			ilGetFunc.Emit(OpCodes.Ldtoken, property.DeclaringType);
			ilGetFunc.Emit(OpCodes.Call, typeof(Type).GetMethod("GetTypeFromHandle", new[] { typeof(RuntimeTypeHandle) }));
			ilGetFunc.Emit(OpCodes.Ldstr, property.Name);
			ilGetFunc.Emit(OpCodes.Call, typeof(Type).GetMethod("GetProperty", new[] { typeof(string) }));

			// IDispatchAccess.GetProperty(this._src, property.Name, indices)
			ilGetFunc.Emit(OpCodes.Ldarg_0);
			ilGetFunc.Emit(OpCodes.Ldfld, srcField);
			ilGetFunc.Emit(OpCodes.Ldstr, property.Name);
			ilGetFunc.Emit(OpCodes.Ldloc_0);
			ilGetFunc.Emit(
				OpCodes.Call,
				GetGetPropertyMethod()
			);

			ilGetFunc.Emit(OpCodes.Callvirt, typeof(IReadValueConverter).GetMethod("Convert", new[] { typeof(PropertyInfo), typeof(object) }));

			if (property.PropertyType.IsValueType)
				ilGetFunc.Emit(OpCodes.Unbox_Any, property.PropertyType);

			ilGetFunc.Emit(OpCodes.Ret);
		}

		private static MethodInfo GetGetPropertyMethod()
		{
			var getPropertyMethod = typeof(IDispatchAccess).GetMethods()
				.Where(m => m.Name == "GetProperty")
				.Where(m => !m.IsGenericMethod)
				.Where(m =>
				{
					var parameters = m.GetParameters();
					if ((parameters == null) || (parameters.Length != 3))
						return false;
					return (
						(parameters[0].ParameterType == typeof(object)) &&
						(parameters[1].ParameterType == typeof(string)) &&
						(parameters[2].ParameterType == typeof(object[]))
					);
				})
				.FirstOrDefault();
			if (getPropertyMethod == null)
				throw new Exception("Failed to locate IDispatchAccess.GetProperty");
			return getPropertyMethod;
		}

		private static void WritePropertySetter(MethodBuilder funcBuilder, FieldBuilder srcField, PropertyInfo property)
		{
			if (funcBuilder == null)
				throw new ArgumentNullException("funcBuilder");
			if (srcField == null)
				throw new ArgumentNullException("srcField");
			if (property == null)
				throw new ArgumentNullException("property");

			var ilSetFunc = funcBuilder.GetILGenerator();

			// Generate: object[] indexParameters
			var indexParameterValues = ilSetFunc.DeclareLocal(typeof(object[]));

			// Generate: indexParameters = new object[x]
			// - The first x number of arguments will be indices and the last argument will be the value to set (for cases where there this
			//   is not an indexed property, x will be zero and the ONLY argument will be the value to set)
			var indexParameters = property.GetIndexParameters();
			var indexParameterTypes = new List<Type>();
			foreach (var indexParameter in indexParameters)
			{
				// These options aren't supported by C# but to be on the safe side we'll ensure they're not enabled some other way
				if (indexParameter.IsOut)
					throw new ArgumentException("Output parameters are not supported");
				if (indexParameter.IsOptional)
					throw new ArgumentException("Optional parameters are not supported");
				if (indexParameter.ParameterType.IsByRef)
					throw new ArgumentException("Ref parameters are not supported");
				indexParameterTypes.Add(indexParameter.ParameterType);
			}
			ilSetFunc.Emit(OpCodes.Ldc_I4, indexParameters.Length);
			ilSetFunc.Emit(OpCodes.Newarr, typeof(Object));
			ilSetFunc.Emit(OpCodes.Stloc_0);
			for (var index = 0; index < indexParameters.Length; index++)
			{
				// Generate: indexParameters[n] = ..;
				var parameter = indexParameters[index];
				ilSetFunc.Emit(OpCodes.Ldloc_0);
				ilSetFunc.Emit(OpCodes.Ldc_I4, index);
				ilSetFunc.Emit(OpCodes.Ldarg, index + 1);
				if (parameter.ParameterType.IsValueType)
					ilSetFunc.Emit(OpCodes.Box, parameter.ParameterType);
				ilSetFunc.Emit(OpCodes.Stelem_Ref);
			}

			// IDispatchAccess.SetProperty(this._src, property.Name, value, indices)
			ilSetFunc.Emit(OpCodes.Ldarg_0);
			ilSetFunc.Emit(OpCodes.Ldfld, srcField);
			ilSetFunc.Emit(OpCodes.Ldstr, property.Name);
			ilSetFunc.Emit(OpCodes.Ldarg, indexParameters.Length + 1);
			ilSetFunc.Emit(OpCodes.Ldloc_0);
			ilSetFunc.Emit(
				OpCodes.Call,
				typeof(IDispatchAccess).GetMethod("SetProperty", new[] { typeof(object), typeof(string), typeof(object), typeof(object[]) })
			);

			ilSetFunc.Emit(OpCodes.Ret);
		}

		private static void generateMethods<T>(TypeBuilder typeBuilder, FieldBuilder srcField, FieldBuilder readValueConverterField, NonNullImmutableList<Type> interfaces)
		{
			if (!typeof(T).IsInterface)
				throw new ArgumentException("typeparam must be an interface type", "targetInterface");
			if (typeBuilder == null)
				throw new ArgumentNullException("typeBuilder");
			if (srcField == null)
				throw new ArgumentNullException("srcField");
			if (readValueConverterField == null)
				throw new ArgumentNullException("readValueConverterField");
			if (!typeof(IReadValueConverter).IsAssignableFrom(readValueConverterField.FieldType))
				throw new ArgumentException("readValueConverterField must be assignable to IReadValueConverter");
			if (interfaces == null)
				throw new ArgumentNullException("interfaces");

			// There could be multiple methods defined with the same signature if they appear in multiple interfaces, this doesn't actually cause
			// any problems so we won't bother doing any work to prevent it happening
			// - While enumerating methods to implement, the get/set methods related to properties will be picked up here and duplicated, but
			//   this also doesn't cause any ill effects so no work is done to prevent it
			foreach (var method in interfaces.SelectMany(i => i.GetMethods()))
			{
				var parameters = method.GetParameters();
				var parameterTypes = new List<Type>();
				foreach (var parameter in parameters)
				{
					if (parameter.IsOut)
						throw new ArgumentException("Output parameters are not supported");
					if (parameter.IsOptional)
						throw new ArgumentException("Optional parameters are not supported");
					if (parameter.ParameterType.IsByRef)
						throw new ArgumentException("Ref parameters are not supported");
					parameterTypes.Add(parameter.ParameterType);
				}
				var funcBuilder = typeBuilder.DefineMethod(
					method.Name,
					MethodAttributes.Public | MethodAttributes.HideBySig | MethodAttributes.NewSlot | MethodAttributes.Virtual | MethodAttributes.Final,
					method.ReturnType,
					parameterTypes.ToArray()
				);
				var ilFunc = funcBuilder.GetILGenerator();

				// Re-use the property getter/setter logic we already have
				var propertyThatThisIsTheGetterFor = method.DeclaringType.GetProperties().FirstOrDefault(prop => prop.GetGetMethod() == method);
				if (propertyThatThisIsTheGetterFor != null)
				{
					WritePropertyGetter(funcBuilder, srcField, readValueConverterField, propertyThatThisIsTheGetterFor);
					continue;
				}
				var propertyThatThisIsTheSetterFor = method.DeclaringType.GetProperties().FirstOrDefault(prop => prop.GetSetMethod() == method);
				if (propertyThatThisIsTheSetterFor != null)
				{
					WritePropertySetter(funcBuilder, srcField, propertyThatThisIsTheSetterFor);
					continue;
				}

				// Generate: object[] args
				var argValues = ilFunc.DeclareLocal(typeof(object[]));

				// Generate: args = new object[x]
				ilFunc.Emit(OpCodes.Ldc_I4, parameters.Length);
				ilFunc.Emit(OpCodes.Newarr, typeof(Object));
				ilFunc.Emit(OpCodes.Stloc_0);
				for (var index = 0; index < parameters.Length; index++)
				{
					// Generate: args[n] = ..;
					var parameter = parameters[index];
					ilFunc.Emit(OpCodes.Ldloc_0);
					ilFunc.Emit(OpCodes.Ldc_I4, index);
					ilFunc.Emit(OpCodes.Ldarg, index + 1);
					if (parameter.ParameterType.IsValueType)
						ilFunc.Emit(OpCodes.Box, parameter.ParameterType);
					ilFunc.Emit(OpCodes.Stelem_Ref);
				}

				// Generate one of either:
				// 1. IDispatchAccess.CallMethod(this._src, method.Name, args)
				// 2. return this._readValueConverter.Convert(
				//  method.DeclaringType.GetMethod(method.Name, {MethodArgTypes})
				//  IDispatchAccess.CallMethod(this._src, method.Name, args)
				// );
				if (!method.ReturnType.Equals(typeof(void)))
				{
					// Generate: Type[] argTypes
					var argTypes = ilFunc.DeclareLocal(typeof(Type[]));

					// Generate: argTypes = new Type[x]
					ilFunc.Emit(OpCodes.Ldc_I4, parameters.Length);
					ilFunc.Emit(OpCodes.Newarr, typeof(Type));
					ilFunc.Emit(OpCodes.Stloc_1);
					for (var index = 0; index < parameters.Length; index++)
					{
						// Generate: argTypes[n] = ..;
						var parameter = parameters[index];
						ilFunc.Emit(OpCodes.Ldloc_1);
						ilFunc.Emit(OpCodes.Ldc_I4, index);
						ilFunc.Emit(OpCodes.Ldtoken, parameters[index].ParameterType);
						ilFunc.Emit(OpCodes.Stelem_Ref);
					}

					// Will call readValueConverter.Convert, passing MethodInfo reference before value
					ilFunc.Emit(OpCodes.Ldarg_0);
					ilFunc.Emit(OpCodes.Ldfld, readValueConverterField);
					ilFunc.Emit(OpCodes.Ldtoken, method.DeclaringType);
					ilFunc.Emit(OpCodes.Call, typeof(Type).GetMethod("GetTypeFromHandle", new[] { typeof(RuntimeTypeHandle) }));
					ilFunc.Emit(OpCodes.Ldstr, method.Name);
					ilFunc.Emit(OpCodes.Ldloc_1);
					ilFunc.Emit(OpCodes.Call, typeof(Type).GetMethod("GetMethod", new[] { typeof(string), typeof(Type[]) }));
				}

				ilFunc.Emit(OpCodes.Ldarg_0);
				ilFunc.Emit(OpCodes.Ldfld, srcField);
				ilFunc.Emit(OpCodes.Ldstr, method.Name);
				ilFunc.Emit(OpCodes.Ldloc_0);
				ilFunc.Emit(OpCodes.Call, GetCallMethodMethod());

				if (!method.ReturnType.Equals(typeof(void)))
					ilFunc.Emit(OpCodes.Callvirt, typeof(IReadValueConverter).GetMethod("Convert", new[] { typeof(MethodInfo), typeof(object) }));

				if (method.ReturnType.Equals(typeof(void)))
					ilFunc.Emit(OpCodes.Pop);
				else if (method.ReturnType.IsValueType)
					ilFunc.Emit(OpCodes.Unbox_Any, method.ReturnType);

				ilFunc.Emit(OpCodes.Ret);
			}
		}

		private static MethodInfo GetCallMethodMethod()
		{
			var callMethodMethod = typeof(IDispatchAccess).GetMethods()
				.Where(m => m.Name == "CallMethod")
				.Where(m => !m.IsGenericMethod)
				.Where(m =>
				{
					var parameters = m.GetParameters();
					if ((parameters == null) || (parameters.Length != 3))
						return false;
					return (
						(parameters[0].ParameterType == typeof(object)) &&
						(parameters[1].ParameterType == typeof(string)) &&
						(parameters[2].ParameterType == typeof(object[]))
					);
				})
				.FirstOrDefault();
			if (callMethodMethod == null)
				throw new Exception("Failed to locate IDispatchAccess.CallMethod");
			return callMethodMethod;
		}

		// ================================================================================================================================
		// IInterfaceApplier IMPLEMENTATION
		// ================================================================================================================================
		private class InterfaceApplier<T> : IInterfaceApplier<T>
		{
			private Func<object, T> _conversion;
			public InterfaceApplier(Func<object, T> conversion)
			{
				if (!typeof(T).IsInterface)
					throw new ArgumentException("Invalid typeparam - must be an interface");
				if (conversion == null)
					throw new ArgumentNullException("conversion");
				_conversion = conversion;
			}

			public Type TargetType
			{
				get { return typeof(T); }
			}

			public T Apply(object src)
			{
				return _conversion(src);
			}

			object IInterfaceApplier.Apply(object src)
			{
				return Apply(src);
			}
		}
	}
}
