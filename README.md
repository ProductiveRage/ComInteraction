# Dynamic COM object wrappers to .Net interfaces

This was written to deal with integrating C# code with various WSC controls (see [Introducing Windows Scripting Components](http://msdn.microsoft.com/en-us/library/07zhfkh8.aspx)) - I wanted to be able to load a component from an arbitrary .wsc file or from a registered component and apply an interface to it which C# code could call it through.

It achieves this by examining a target interface (and its hierarchy, if it implements any other interfaces) and generating a new class by emitting IL that will take an object reference and hook up the required properties and methods to IDispatch calls against the source object, effectively applying the interface over the source reference. The wrapping class can optionally be generated with ComVisible attributes such that it can be passed into other WSC controls (or other COM objects).

eg. Here's a simple WSC control -

    <?xml version="1.0" ?>
    <?component error="false" debug="false" ?>
    <package>
      <component id="TestControl">
      <registration progid="COMInteraction.TestControl" description="Test WSC" version="1" />
          <public>
              <property name="InterfaceVersion" />
              <property name="Application" />
              <method name="Init" />
              <method name="GetRenderDependencies" />
              <method name="Render">
                  <parameter name="pO" />
              </method>
              <method name="Dispose" />
          </public>
          <script language="VBScript">
          <![CDATA[
              Option Explicit

              Dim InterfaceVersion, Application

              InterfaceVersion = 1
              Set Application = new WscApplication
    
              Public Function Init()
              End Function

              Public Function GetRenderDependencies()
                  GetRenderDependencies = "Common"
              End Function

              Public Function Render(ByVal writer)
                  writer.Write "Whoop! " & TypeName(writer)
              End Function

              Public Function Dispose()
              End Function

              Class WscApplication
                  Public Property Get Name
                      Name = "WscApplication"
                  End Property
              End Class
          ]]>
          </script>
      </component>
    </package>

that I want to apply the following interface to..

    public interface IControl
    {
        short InterfaceVersion { get; }
        IApplication Application { get; set; }
        void Init();
        string GetRenderDependencies();
        void Render(IOutputWriter writer);
        void Dispose();
    }

    public interface IApplication
    {
        string Name { get; }
    }

    public interface IOutputWriter
    {
        void Write(string content);
    }

I can do this using the following code:

    var interfaceApplierFactory = new IDispatchInterfaceApplierFactory(
        "DynamicAssembly",
        ComVisibilityOptions.Visible
    );
    var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<IControl>(
        new CachedReadValueConverter(interfaceApplierFactory)
    );
    var obj = (new COMObjectLoader()).LoadFromScriptFile(filenameOfWscControl);
    var objWrapped = interfaceApplier.Apply(obj);

and then I can access the WSC control with:

    Console.WriteLine("Application.Name: " + objWrapped.Application.Name);
    Console.WriteLine("InterfaceVersion: " + objWrapped.InterfaceVersion);
    objWrapped.Init();
    Console.WriteLine(objWrapped.GetRenderDependencies());
    var writer = new COMOutputWriter();
    objWrapped.Render(writer);
    objWrapped.Dispose();

Note that the Application property of the generated IControl is wrapped in a translation reference such that the IApplication interface is applied correctly to the objWrapped.Application reference.

## Type Libraries and Dynamic

I'm sure that we could have got some of the way here by ensuring that we generated type libraries for any WSC control and then importing them as references into a .Net project but I wanted to see if I could provide the flexibility of pointing at any WSC file and applying any appropriate interface.

I also looked into the C# 4.0 "dynamic" keyword to see if that covered much of the same ground - like if you could just apply an interface to a dynamic object reference and have it all work by magic somehow. I don't think you can. It should be possible to generate IL to use a dynamic object reference as the wrapped value instead of an object whose properties and methods are called by reflection - I had a play around and it looks like this would improve the performance of the property and method access but at the cost of more complicated IL generation. I'm not sure if the argument that we're already paying a penalty to cross COM boundaries is _for_ investigating improved performance using dynamic or _against_ it! (One advantage that this code has is that it works with .Net 3.5 but it also doesn't have the full range of integration options that dynamic has, dynamic can interact through many more mechanisms that just IDispatch and reflection).

## Update (20th February 2013): .Net types and reflection

I've updated this project with alternate Interface Applier Factory classes; ReflectionInterfaceApplierFactory, which is intended for wrapping interfaces around .Net types (so reflection can be used rather than trying to go through IDispatch) and the CombinedInterfaceApplierFactory which will try to use which ever of the two is appropriate for the source data (or use neither if no conversion is required).

    var dynamicAssemblyName = "DynamicAssembly";
    var interfaceApplierFactory = 
        new ReflectionInterfaceApplierFactory(
            dynamicAssemblyName,
            ComVisibilityOptions.Visible
        ),
        new IDispatchInterfaceApplierFactory(
            dynamicAssemblyName,
            ComVisibilityOptions.Visible
        )
    );

The appropriate conversion method will be used for each conversion, returned types and their properties are recursively wrapped as necessary as well (as with the IApplication wrapping in the original example). No conversion will be applied if a return type already implements the interface that is requested of it. If a conversion _is_ required that can't be dealt with by reflection or IDispatch then an exception will be raised.