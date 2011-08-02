using System;

namespace Tester
{
    public interface IOutputWriter
    {
        void Append(string content);
        void AppendLine(string content);
        void Write(string content);
        void WriteLine(string content);
    }
}
