using System;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace NetOffice.Word
{
    /// <summary>
    /// Microsoft Word application object.
    /// </summary>
    public class Application : _Application
    {
        [SupportedOSPlatform("windows")]
        public Application(object instance) : base(instance)
        {
        }
    }
}
