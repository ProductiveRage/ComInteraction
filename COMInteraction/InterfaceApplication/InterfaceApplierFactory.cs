using System;
using System.Collections.Generic;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Threading;
using COMInteraction.InterfaceApplication.ReadValueConverters;

namespace COMInteraction.InterfaceApplication
{
    /// <summary>
    /// Return a delegate that will take an object reference and return an implementation of a specified interface - any calls to properties or methods of that interface
    /// are passed straight through to the wrapped reference using Reflection. If any of the interface's properties or methods are not available on the wrapped reference
    /// then an exception will be thrown when requested. Any interfaces that the target interface implements will be accounted for. This can be used to wrap POCO objects
    /// or COM objects, for example. Note: Only handling interfaces, not classes, means there are less things to consider - there are no protected set methods or fields,
    /// for example.
    /// </summary>
    public class InterfaceApplierFactory : IInterfaceApplierFactory
    {
        // ================================================================================================================================
        // CLASS INITIALISATION
        // ================================================================================================================================
        private string _assemblyName;
        private bool _createComVisibleClasses;
        private Lazy<ModuleBuilder> _moduleBuilder;
        public InterfaceApplierFactory(string assemblyName, ComVisibility comVisibilityOfClasses)
        {
            assemblyName = (assemblyName ?? "").Trim();
            if (assemblyName == "")
                throw new ArgumentException("Null or empty assemblyName specified");
            if (!Enum.IsDefined(typeof(ComVisibility), comVisibilityOfClasses))
                throw new ArgumentOutOfRangeException("comVisibilityOfClasses");
            
            _assemblyName = assemblyName;
            _createComVisibleClasses = (comVisibilityOfClasses == ComVisibility.Visible);
            _moduleBuilder = new Lazy<ModuleBuilder>(
                () =>
                {
                    var assemblyBuilder = Thread.GetDomain().DefineDynamicAssembly(
                        new AssemblyName(_assemblyName),
                        AssemblyBuilderAccess.Run
                    );
                    return assemblyBuilder.DefineDynamicModule(
                        assemblyBuilder.GetName().Name,
                        false
                    );
                },
                true // isThreadSafe
            );
        }

        public enum ComVisibility
        {
            Visible,
            NotVisible
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

            var typeName = "InterfaceApplier" + Guid.NewGuid().ToString("N");
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
            var interfaces = new List<Type>();
            buildInterfaceInheritanceList(typeof(T), interfaces);

            // ================================================================================================
            // Constructor and private "src" and "readValueConverter" fields
            // ================================================================================================
            var srcField = typeBuilder.DefineField("_src", typeof(object), FieldAttributes.Private);
            var readValueConverterField = typeBuilder.DefineField("_readValueConverter", typeof(IReadValueConverter), FieldAttributes.Private);
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

            // ================================================================================================
            // Properties (since we're only dealing with interfaces, we need only conside the CanRead and
            // CanWrite values on the properties, nothing more complicated)
            // ================================================================================================
            // There could be multiple properties defined with different types if they appear in multiple interfaces, this doesn't actually cause
            // any problems so we won't bother doing any work to prevent it happening
            var properties = new List<PropertyInfo>();
            foreach (var entry in interfaces)
                properties.AddRange(entry.GetProperties());

            foreach (var property in properties)
            {
                var methodInfoInvokeMember = typeof(Type).GetMethod(
                    "InvokeMember",
                    new[]
                    {
                        typeof(string),
                        typeof(BindingFlags),
                        typeof(Binder),
                        typeof(object),
                        typeof(object[])
                    }
                );

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

                    // Generate: return this._readValueConverter.ConvertPropertyValue(
                    //  typeof(T).GetProperty(property.Name)
                    //  this._src.GetType().InvokeMember(property.Name, BindingFlags.GetProperty, null, _src, null)
                    // );
                    var ilGetFunc = getFuncBuilder.GetILGenerator();
                    
                    ilGetFunc.Emit(OpCodes.Ldarg_0);
                    ilGetFunc.Emit(OpCodes.Ldfld, readValueConverterField);
                    ilGetFunc.Emit(OpCodes.Ldtoken, typeof(T));
                    ilGetFunc.Emit(OpCodes.Call, typeof(Type).GetMethod("GetTypeFromHandle", new[] { typeof(RuntimeTypeHandle) }));
                    ilGetFunc.Emit(OpCodes.Ldstr, property.Name);
                    ilGetFunc.Emit(OpCodes.Call, typeof(Type).GetMethod("GetProperty", new[] { typeof(string) }));

                    ilGetFunc.Emit(OpCodes.Ldarg_0);
                    ilGetFunc.Emit(OpCodes.Ldfld, srcField);
                    ilGetFunc.Emit(OpCodes.Callvirt, typeof(Type).GetMethod("GetType", Type.EmptyTypes));
                    ilGetFunc.Emit(OpCodes.Ldstr, property.Name);
                    ilGetFunc.Emit(OpCodes.Ldc_I4, (int)BindingFlags.GetProperty);
                    ilGetFunc.Emit(OpCodes.Ldnull);
                    ilGetFunc.Emit(OpCodes.Ldarg_0);
                    ilGetFunc.Emit(OpCodes.Ldfld, srcField);
                    ilGetFunc.Emit(OpCodes.Ldnull);
                    ilGetFunc.Emit(OpCodes.Callvirt, methodInfoInvokeMember);
                    ilGetFunc.Emit(OpCodes.Callvirt, typeof(IReadValueConverter).GetMethod("Convert", new[] { typeof(PropertyInfo), typeof(object) }));

                    if (property.PropertyType.IsValueType)
                        ilGetFunc.Emit(OpCodes.Unbox_Any, property.PropertyType);

                    ilGetFunc.Emit(OpCodes.Ret);
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
                    var valueParameter = setFuncBuilder.DefineParameter(1, ParameterAttributes.None, "value");
                    var ilSetFunc = setFuncBuilder.GetILGenerator();
                    // Generate: this._src.GetType().InvokeMember(property.Name, BindingFlags.SetProperty, null, _src, new object[1] { value });
                    var argValues = ilSetFunc.DeclareLocal(typeof(object[])); // Need to declare assignment of local array to pass to InvokeMember
                    ilSetFunc.Emit(OpCodes.Ldarg_0);
                    ilSetFunc.Emit(OpCodes.Ldfld, srcField);
                    ilSetFunc.Emit(OpCodes.Callvirt, typeof(Type).GetMethod("GetType", Type.EmptyTypes));
                    ilSetFunc.Emit(OpCodes.Ldstr, property.Name);
                    ilSetFunc.Emit(OpCodes.Ldc_I4, (int)BindingFlags.SetProperty);
                    ilSetFunc.Emit(OpCodes.Ldnull);
                    ilSetFunc.Emit(OpCodes.Ldarg_0);
                    ilSetFunc.Emit(OpCodes.Ldfld, srcField);
                    ilSetFunc.Emit(OpCodes.Ldc_I4_1);
                    ilSetFunc.Emit(OpCodes.Newarr, typeof(Object));
                    ilSetFunc.Emit(OpCodes.Stloc_0);
                    ilSetFunc.Emit(OpCodes.Ldloc_0);
                    ilSetFunc.Emit(OpCodes.Ldc_I4_0);
                    ilSetFunc.Emit(OpCodes.Ldarg_1);
                    if (property.PropertyType.IsValueType)
                        ilSetFunc.Emit(OpCodes.Box, property.PropertyType);
                    ilSetFunc.Emit(OpCodes.Stelem_Ref);
                    ilSetFunc.Emit(OpCodes.Ldloc_0);
                    ilSetFunc.Emit(OpCodes.Callvirt, methodInfoInvokeMember);
                    ilSetFunc.Emit(OpCodes.Pop);

                    ilSetFunc.Emit(OpCodes.Ret);
                    propBuilder.SetSetMethod(setFuncBuilder);
                }
            }

            // ================================================================================================
            // Functions
            // ================================================================================================
            // There could be multiple methods defined with the same signature if they appear in multiple interfaces, this doesn't actually cause
            // any problems so we won't bother doing any work to prevent it happening
            // - While enumerating methods to implement, the get/set methods related to properties will be picked up here and duplicated, but
            //   this also doesn't cause any ill effects so no work is done to prevent it
            var methods = new List<MethodInfo>();
            foreach (var entry in interfaces)
                methods.AddRange(entry.GetMethods());

            foreach (var method in methods)
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
                // 1. _src.GetType().InvokeMember(method.Name, BindingFlags.InvokeMethod, null, _src, args);
                // 2. return this._readValueConverter.ConvertPropertyValue(
                //  typeof(T).GetMethod(method.Name, {MethodArgTypes})
                //  this._src.GetType().InvokeMember(property.Name, BindingFlags.GetProperty, null, _src, null)
                // );
                var methodInfoInvokeMember = typeof(Type).GetMethod(
                    "InvokeMember",
                    new[]
                    {
                        typeof(string),
                        typeof(BindingFlags),
                        typeof(Binder),
                        typeof(object),
                        typeof(object[])
                    }
                );

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
                        ilFunc.Emit(OpCodes.Ldtoken, parameters[index].GetType());
                        ilFunc.Emit(OpCodes.Stelem_Ref);
                    }

                    // Will call readValueConverter.Convert, passing MethodInfo reference before value
                    ilFunc.Emit(OpCodes.Ldarg_0);
                    ilFunc.Emit(OpCodes.Ldfld, readValueConverterField);
                    ilFunc.Emit(OpCodes.Ldtoken, typeof(T));
                    ilFunc.Emit(OpCodes.Call, typeof(Type).GetMethod("GetTypeFromHandle", new[] { typeof(RuntimeTypeHandle) }));
                    ilFunc.Emit(OpCodes.Ldstr, method.Name);
                    ilFunc.Emit(OpCodes.Ldloc_1);
                    ilFunc.Emit(OpCodes.Call, typeof(Type).GetMethod("GetMethod", new[] { typeof(string), typeof(Type[]) }));
                }

                ilFunc.Emit(OpCodes.Ldarg_0);
                ilFunc.Emit(OpCodes.Ldfld, srcField);
                ilFunc.Emit(OpCodes.Callvirt, typeof(Type).GetMethod("GetType", Type.EmptyTypes));
                ilFunc.Emit(OpCodes.Ldstr, method.Name);
                ilFunc.Emit(OpCodes.Ldc_I4, (int)BindingFlags.InvokeMethod);
                ilFunc.Emit(OpCodes.Ldnull);
                ilFunc.Emit(OpCodes.Ldarg_0);
                ilFunc.Emit(OpCodes.Ldfld, srcField);
                ilFunc.Emit(OpCodes.Ldloc_0);
                ilFunc.Emit(OpCodes.Callvirt, methodInfoInvokeMember);

                if (!method.ReturnType.Equals(typeof(void)))
                    ilFunc.Emit(OpCodes.Callvirt, typeof(IReadValueConverter).GetMethod("Convert", new[] { typeof(MethodInfo), typeof(object) }));

                if (method.ReturnType.Equals(typeof(void)))
                    ilFunc.Emit(OpCodes.Pop);
                else if (method.ReturnType.IsValueType)
                    ilFunc.Emit(OpCodes.Unbox_Any, method.ReturnType);
                
                ilFunc.Emit(OpCodes.Ret);
            }
            
            return new InterfaceApplier<T>(
                src => (T)Activator.CreateInstance(
                    typeBuilder.CreateType(),
                    src,
                    readValueConverter
                )
            );
        }

        private void buildInterfaceInheritanceList(Type targetInterface, List<Type> types)
        {
            if (targetInterface == null)
                throw new ArgumentNullException("targetInterface");
            if (!targetInterface.IsInterface)
                throw new ArgumentException("targetInterface must be an interface type", "targetInterface");
            if (types == null)
                throw new ArgumentNullException("types");
            if (!types.Contains(targetInterface))
                types.Add(targetInterface);
            foreach (var inheritedInterface in targetInterface.GetInterfaces())
            {
                if (!types.Contains(inheritedInterface))
                {
                    types.Add(inheritedInterface);
                    buildInterfaceInheritanceList(inheritedInterface, types);
                }
            }
        }

        // ================================================================================================================================
        // CHILD CLASSES
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

        private sealed class AssemblyType
        {
            private string _assemblyName;
            private Type _targetType;
            public AssemblyType(string assemblyName, Type targetType)
            {
                assemblyName = (assemblyName ?? "").Trim();
                if (assemblyName == "")
                    throw new ArgumentException("Null or empty assemblyName specified");
                if (targetType == null)
                    throw new ArgumentNullException("targetType");
                _assemblyName = assemblyName;
                _targetType = targetType;
            }
            public string AssemblyName { get { return _assemblyName; } }
            public Type TargetType { get { return _targetType; } }
            public override bool Equals(object obj)
            {
                var objAssemblyType = obj as AssemblyType;
                if (objAssemblyType == null)
                    return false;
                return ((objAssemblyType.AssemblyName == this.AssemblyName) && (objAssemblyType.TargetType == this.TargetType));
            }
            public override int GetHashCode()
            {
                return String.Format("{0}_{1}_{2}", this.GetType().ToString(), this.TargetType.ToString(), this.AssemblyName).GetHashCode();
            }
        }
    }
}
