using System;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace NetOffice.PowerPoint
{
    //[TypeLibType(4288)]
    [Guid("91493442-5a91-11cf-8700-00aa0060263b")]
    public class _Application : Dispatcher
    {
        [SupportedOSPlatform("windows")]
        public _Application(object instance) : base(instance)
        {
        }

        public _Application(IntPtr unknownPtr) : base(unknownPtr)
        {
        }

        public string Name
        {
            get { return this.InvokePropertyGet<string>(0); }
        }

        public string Caption
        {
            get { return this.InvokePropertyGet<string>(2009); }
            set { this.InvokePropertySet<string>(2009, value); }
        }
    }
}
