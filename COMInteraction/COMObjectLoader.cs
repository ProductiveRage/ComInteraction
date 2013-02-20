using System;
using Microsoft.VisualBasic;

namespace COMInteraction
{
    public class COMObjectLoader
    {
        /// <summary>
        /// Instantiate a new COMObject wrapper given the ProgId
        /// </summary>
        public object LoadFromProgId(string progId)
        {
            if ((progId ?? "").Trim() == "")
                throw new ArgumentException("Null or empty progId specified");
            return Activator.CreateInstance(
                Type.GetTypeFromProgID(progId, true) // Pass true for throwOnError
            );
        }

        /// <summary>
        /// Instantiate a new COMObject wrapper given the ProgId
        /// </summary>
        public object LoadFromScriptFile(string filename)
        {
			if ((filename ?? "").Trim() == "")
                throw new ArgumentException("Null or empty filename specified");
            return Interaction.GetObject("script:" + filename, null);
        }
    }
}
