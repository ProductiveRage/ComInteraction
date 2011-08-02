# Dynamic COM object wrappers to .Net interfaces

This was written to deal with integrating C# code with various WSC controls (see [Introducing Windows Scripting Components](http://msdn.microsoft.com/en-us/library/07zhfkh8.aspx)) - I wanted to be able to load a component from an arbitrary .wsc file or from a registered component and apply an interface to it which C# code could call it through.

It achieves this by examining a target interface (and its hierarchy, if it implements any other interfaces) and generating a new class by emitting IL that will take an object reference and hook up the required properties and methods to reflection calls against the source object, effectively applying the interface over the source reference. The wrapping class can optionally be generated with ComVisible attributes such that it can be passed into other WSC controls (or other COM objects).

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

    var interfaceApplierFactory = new InterfaceApplierFactory(
        "DynamicAssembly",
        InterfaceApplierFactory.ComVisibility.Visible
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

I also looked into the C# 4.0 "dynamic" keyword to see if that covered much of the same ground - like if you could just apply an interface to a dynamic object reference and have it all work by magic somehow. I don't think you can. But there may be a way to generate IL to use a dynamic object reference as the wrapped value instead of an object whose properties and methods are called by reflection - I don't know enough about "dynamic" yet to know if this would bring any benefit.