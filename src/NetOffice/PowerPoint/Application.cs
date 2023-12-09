using System;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace NetOffice.PowerPoint
{
    [Guid("91493442-5A91-11CF-8700-00AA0060263B")]
    public class Application : _Application
    {
        [SupportedOSPlatform("windows")]
        public Application(object instance) : base(instance)
        {
        }

        public Application(IntPtr unknownPtr) : base(unknownPtr)
        {
        }
    }
}
