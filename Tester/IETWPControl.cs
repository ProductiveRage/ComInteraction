using System;

namespace Tester
{
    public interface IETWPControl
    {
        Int16 InterfaceVersion { get; }

        IApplication Application { get; set; }
        object Cache { get; set; }
        object Config { get; set; }
        object DMSCache { get; set; }
        object Request { get; set; }
        object Response { get; set; }
        object Server { get; set; }
        object Session { get; set; }

        object Context { get; set; }
        object Page { get; set; }
        object Params { get; set; }
        object PlaceHolder { get; set; }

        void Init();
        void Load();
        void PreRender();
        Int16 GetOutputCacheType();
        string GetRenderDependencies();
        void Render(object pO);
        void RenderPartial(object pO);
        void Dispose();
    }
}
