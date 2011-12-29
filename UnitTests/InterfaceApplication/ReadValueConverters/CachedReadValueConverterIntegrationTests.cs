using COMInteraction.InterfaceApplication;
using COMInteraction.InterfaceApplication.ReadValueConverters;
using Xunit;

namespace UnitTests.InterfaceApplication.ReadValueConverters
{
    public class CachedReadValueConverterIntegrationTests
    {
        /// <summary>
        /// The CachedReadValueConverter previously had a callChain mechanism to ensure that the same interface wasn't to be applied to na object as part of the process of applying
        /// that same interface to an object further up the interface-applying chain, the thinking being that this would end up in an infinite loop that would cause a stack overflow.
        /// However, this does not appear to be the case as demonstrated here, so the callChain testing has been removed from CachedReadValueConverter.
        /// </summary>
        [Fact]
        public void RecursiveInterfacesDoNoCauseStackOverflow()
        {
            var interfaceApplierFactory = new InterfaceApplierFactory(
                "DynamicAssembly",
                InterfaceApplierFactory.ComVisibility.Visible
            );

            var n1 = new Node() { Name = "Node1" };
            var n2 = new Node() { Name = "Node2" };
            var n3 = new Node() { Name = "Node3" };

            n1.Hierarchy = new Node.HierarchyData() { Self = n1, Next = n2 };

            n2.Hierarchy = new Node.HierarchyData()
            {
                Previous = n1,
                Self = n2,
                Next = n3
            };

            n3.Hierarchy = new Node.HierarchyData() { Previous = n2, Self = n3 };

            var interfaceApplier = interfaceApplierFactory.GenerateInterfaceApplier<INode>(
                new CachedReadValueConverter(interfaceApplierFactory)
            );
            var n2Wrapped = interfaceApplier.Apply(n2);

            // Navigating right around the hierarchy chain here to try to ensure that if an overflow is possible that it would occur here
            Assert.Equal(
                "Node2",
                n2Wrapped.Hierarchy.Previous.Hierarchy.Next.Hierarchy.Next.Hierarchy.Previous.Name
            );
        }

        public class Node
        {
            public string Name { get; set; }
            public HierarchyData Hierarchy { get; set; }
            public class HierarchyData
            {
                public Node Previous { get; set; }
                public Node Self { get; set; }
                public Node Next { get; set; }
            }
        }

        public interface INode
        {
            string Name { get; set; }
            INodeHierarchy Hierarchy { get; set; }
        }

        public interface INodeHierarchy
        {
            INode Previous { get; set; }
            INode Self { get; set; }
            INode Next { get; set; }
        }
    }
}
