using System;
using System.Reflection;
using COMInteraction.InterfaceApplication;
using COMInteraction.InterfaceApplication.ReadValueConverters;
using UnitTests.InterfaceApplication.Helpers;
using Xunit;

namespace UnitTests.InterfaceApplication
{
    public class InterfaceApplierFactoryTests
    {
        // ================================================================================================================================
        // TESTS: String Properties
        // ================================================================================================================================
        [Fact]
        public void RetrieveStringProperty_NoInheritance()
        {
            var interfaceApplierFactory = new ReflectionInterfaceApplierFactory("InterfaceApplierFactoryTests", ComVisibilityOptions.NotVisible);
            var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<INamedReadOnly>(
                new ActionlessReadValueConverter()
            );
            var src = new ReadOnlyNamedClass1("name");
            var srcWrapped = interfaceApplier.Apply(src);
            Assert.Equal("name", srcWrapped.Name);
        }

        [Fact]
        public void RetrieveStringProperty_PropertyOnInheritedInterface()
        {
            var interfaceApplierFactory = new ReflectionInterfaceApplierFactory("InterfaceApplierFactoryTests", ComVisibilityOptions.NotVisible);
            var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<IPerson>(
                new ActionlessReadValueConverter()
            );
            var src = new ReadOnlyNamedClass1("name");
            var srcWrapped = interfaceApplier.Apply(src);
            Assert.Equal("name", srcWrapped.Name);
        }

        [Fact]
        public void RetrieveStringPropertyWhenWrappedObjectDoesNotHavePropertyWillFail()
        {
            var interfaceApplierFactory = new ReflectionInterfaceApplierFactory("InterfaceApplierFactoryTests", ComVisibilityOptions.NotVisible);
            var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<INamedReadOnly>(
                new ActionlessReadValueConverter()
            );
            var src = new object();
            var srcWrapped = interfaceApplier.Apply(src);
            Assert.Throws<MissingMethodException>(() =>
            {
                Console.WriteLine(srcWrapped.Name);
            });
        }

        // ================================================================================================================================
        // TESTS: Int Properties - will require boxing of value to return as int
        // ================================================================================================================================
        [Fact]
        public void RetrieveIntProperty_NoInheritance()
        {
            var interfaceApplierFactory = new ReflectionInterfaceApplierFactory("InterfaceApplierFactoryTests", ComVisibilityOptions.NotVisible);
            var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<IAgedInt32ReadOnly>(
                new ActionlessReadValueConverter()
            );
            var src = new ReadOnlyAgedInt32Class1(29);
            var srcWrapped = interfaceApplier.Apply(src);
            Assert.Equal(29, srcWrapped.Age);
        }

        [Fact]
        public void RetrieveInt32ValueAsInt16PropertyWillFail()
        {
            var interfaceApplierFactory = new ReflectionInterfaceApplierFactory("InterfaceApplierFactoryTests", ComVisibilityOptions.NotVisible);
            var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<IAgedInt16ReadOnly>(
                new ActionlessReadValueConverter()
            );
            var src = new ReadOnlyAgedInt32Class1(29);
            var srcWrapped = interfaceApplier.Apply(src);
            Assert.Throws<InvalidCastException>(() =>
            {
                Console.WriteLine(srcWrapped.Age);
            });
        }

        [Fact]
        public void RetrieveInt16ValueAsInt32PropertyWillFail()
        {
            var interfaceApplierFactory = new ReflectionInterfaceApplierFactory("InterfaceApplierFactoryTests", ComVisibilityOptions.NotVisible);
            var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<IAgedInt32ReadOnly>(
                new ActionlessReadValueConverter()
            );
            var src = new ReadOnlyAgedInt16Class1(29);
            var srcWrapped = interfaceApplier.Apply(src);
            Assert.Throws<InvalidCastException>(() =>
            {
                Console.WriteLine(srcWrapped.Age);
            });
        }

        [Fact]
        public void RetrieveInt16ValueAsInt32_UsingReadValueConverter()
        {
            var interfaceApplierFactory = new ReflectionInterfaceApplierFactory("InterfaceApplierFactoryTests", ComVisibilityOptions.NotVisible);
            var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<IAgedInt32ReadOnly>(
                new ToInt32PropertyReadValueConverter()
            );
            var src = new ReadOnlyAgedInt16Class1(29);
            var srcWrapped = interfaceApplier.Apply(src);
            Assert.Equal(29, srcWrapped.Age);
        }

        private class ToInt32PropertyReadValueConverter : IReadValueConverter
        {
            public object Convert(PropertyInfo property, object value)
            {
                return System.Convert.ToInt32(value);
            }
            public object Convert(MethodInfo method, object value)
            {
                throw new NotImplementedException();
            }
        }

        // ================================================================================================================================
        // TESTS: Inaccessible Interfaces
        // ================================================================================================================================
        [Fact]
        public void WrappingToInaccessibleInterfaceWillFail()
        {
            var interfaceApplierFactory = new ReflectionInterfaceApplierFactory("InterfaceApplierFactoryTests", ComVisibilityOptions.NotVisible);
            var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<IPrivateNamedReadOnly>(
                new ActionlessReadValueConverter()
            );
            var src = new ReadOnlyNamedClass1("name");
            Assert.Throws<TypeLoadException>(() =>
            {
                interfaceApplier.Apply(src);
            });
        }

        // ================================================================================================================================
        // HELPERS: Target Interfaces
        // ================================================================================================================================
        public interface INamedReadOnly
        {
            string Name { get; }
        }

        public interface IAgedInt32ReadOnly
        {
            int Age { get; }
        }

        public interface IAgedInt16ReadOnly
        {
            short Age { get; }
        }

        public interface IPerson : INamedReadOnly { }

        private interface IPrivateNamedReadOnly : INamedReadOnly { }

        // ================================================================================================================================
        // HELPERS: Target Interface Implementations
        // ================================================================================================================================
        private class ReadOnlyNamedClass1
        {
            public ReadOnlyNamedClass1(string name)
            {
                this.Name = name;
            }
            public string Name { get; private set; }
        }

        private class ReadOnlyAgedInt32Class1
        {
            public ReadOnlyAgedInt32Class1(int age)
            {
                this.Age = age;
            }
            public int Age { get; private set; }
        }

        private class ReadOnlyAgedInt16Class1
        {
            public ReadOnlyAgedInt16Class1(short age)
            {
                this.Age = age;
            }
            public short Age { get; private set; }
        }

        private class ReadWriteNamedClass1
        {
            public string Name { get; set; }
        }
    }
}
