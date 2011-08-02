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
            var interfaceApplierFactory = new InterfaceApplierFactory(
                "DynamicAssembly",
                InterfaceApplierFactory.ComVisibility.Visible
            );

            var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<IETWPControl>(
                new CachedReadValueConverter(interfaceApplierFactory)
            );
            var obj = (new COMObjectLoader()).LoadFromScriptFile(
                Path.Combine(Path.GetDirectoryName(typeof(Program).Assembly.Location), "Test.wsc")
            );

            var objWrapped = interfaceApplier.Apply(obj);

            var app = objWrapped.Application;
            var appName = app.Name;

            Console.WriteLine(objWrapped.Application.Name);

            var version = objWrapped.InterfaceVersion;
            objWrapped.Init();
            objWrapped.Load();
            objWrapped.PreRender();
            var dependencies = objWrapped.GetRenderDependencies();
            var outputCacheType = objWrapped.GetOutputCacheType();
            var outputRender = new COMOutputWriter();
            objWrapped.Render(outputRender);
            var outputRenderPartial = new COMOutputWriter();
            objWrapped.RenderPartial(outputRenderPartial);
            objWrapped.Dispose();
        }
    }
}
