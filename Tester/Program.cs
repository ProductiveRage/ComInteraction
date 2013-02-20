using System;
using System.IO;
using COMInteraction;
using COMInteraction.InterfaceApplication;
using COMInteraction.InterfaceApplication.ReadValueConverters;

namespace Tester
{
	class Program
	{
		static void Main(string[] args)
		{
			DemonstrateIDispatchInterfaceApplication();
			DemonstrateReflectionInterfaceApplication();
		}

		private static void DemonstrateIDispatchInterfaceApplication()
		{
			var interfaceApplierFactory = new IDispatchInterfaceApplierFactory(
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

		private static void DemonstrateReflectionInterfaceApplication()
		{
			var interfaceApplierFactory = new ReflectionInterfaceApplierFactory(
				"DynamicAssembly",
				ComVisibilityOptions.Visible
			);

			var employeeInterfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<IEmployee>(
				new CachedReadValueConverter(interfaceApplierFactory)
			);
			var employee = employeeInterfaceApplier.Apply(
				new Employee("Jim", new Employee.EmployeeRole("Test Pilot"))
			);
		}

		public class Employee
		{
			public Employee(string name, EmployeeRole role)
			{
				if (string.IsNullOrWhiteSpace(name))
					throw new ArgumentException("Null/blank name specified");
				if (role == null)
					throw new ArgumentNullException("role");

				Name = name;
				Role = role;
			}

			public string Name { get; private set; }
			public EmployeeRole Role { get; private set; }

			public class EmployeeRole
			{
				public EmployeeRole(string title)
				{
					if (string.IsNullOrWhiteSpace(title))
						throw new ArgumentException("Null/blank title specified");

					Title = title.Trim();
				}

				public string Title { get; private set; }
			}
		}
	}
}
