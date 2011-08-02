using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Tester
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class COMOutputWriter : IOutputWriter
    {
        private StringBuilder _content;
        public COMOutputWriter()
        {
            _content = new StringBuilder();
        }

        public void Append(string content)
        {
            if (content != null)
                _content.Append(content);
        }

        public void AppendLine(string content)
        {
            if (content != null)
                _content.AppendLine(content);
        }

        public void Write(string content)
        {
            Append(content);
        }

        public void WriteLine(string content)
        {
            AppendLine(content);
        }
    }
}
