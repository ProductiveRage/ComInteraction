using System;
using System.IO;
using System.Runtime.InteropServices;
using COMInteraction;
using COMInteraction.InterfaceApplication;
using COMInteraction.InterfaceApplication.ReadValueConverters;

namespace Tester
{
    class Program
    {
        static void Main(string[] args)
        {
            var interfaceApplierFactory = new ReflectionInterfaceApplierFactory(
                "DynamicAssembly",
                ComVisibilityOptions.Visible
            );

            var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<IControl>(
                new CachedReadValueConverter(interfaceApplierFactory)
            );
            var obj = (new COMObjectLoader()).LoadFromScriptFile(
                Path.Combine(Path.GetDirectoryName(typeof(Program).Assembly.Location), "Test.wsc")
            );

            var objWrapped = interfaceApplier.Apply(obj);
            Console.WriteLine("Application.Name: " + objWrapped.Application.Name);
            Console.WriteLine("InterfaceVersion: " + objWrapped.InterfaceVersion);
            objWrapped.Init();
            Console.WriteLine(objWrapped.GetRenderDependencies());
            var writer = new COMOutputWriter();
            objWrapped.Render(writer);
            objWrapped.Dispose();
        }
    }
}
