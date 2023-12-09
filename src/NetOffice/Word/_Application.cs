using System;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace NetOffice.Word
{
    public class _Application : Dispatcher
    {
        [SupportedOSPlatform("windows")]
        public _Application(object instance) : base(instance)
        {
        }

        public string Name
        {
            get { return this.InvokePropertyGet<string>(0); }
        }
    }
}
