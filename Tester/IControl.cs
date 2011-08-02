using System;

namespace Tester
{
    public interface IControl
    {
        short InterfaceVersion { get; }

        IApplication Application { get; set; }

        void Init();
        string GetRenderDependencies();
        void Render(IOutputWriter writer);
        void Dispose();
    }
}
