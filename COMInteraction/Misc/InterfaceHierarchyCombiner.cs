using System;
using System.Collections.Generic;

namespace COMInteraction.Misc
{
    public class InterfaceHierarchyCombiner
    {
        private Type _targetInterface;
        private NonNullImmutableList<Type> _interfaces;
        public InterfaceHierarchyCombiner(Type targetInterface)
        {
            if (targetInterface == null)
                throw new ArgumentNullException("targetInterface");
            if (!targetInterface.IsInterface)
                throw new ArgumentException("targetInterface must be an interface type", "targetInterface");

            var interfaces = new List<Type>();
            buildInterfaceInheritanceList(targetInterface, interfaces);
            _interfaces = new NonNullImmutableList<Type>(interfaces);
            _targetInterface = targetInterface;
        }

        private static void buildInterfaceInheritanceList(Type targetInterface, List<Type> types)
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

        public Type TargetInterface
        {
            get { return _targetInterface; }
        }
        
        public NonNullImmutableList<Type> Interfaces
        {
            get { return _interfaces; }
        }
    }
}
