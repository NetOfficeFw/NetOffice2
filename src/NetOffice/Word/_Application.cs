using System;
using System.Runtime.InteropServices;

namespace NetOffice.Word
{
    public class _Application : Dispatcher
    {
        public _Application(object instance) : base(instance)
        {
        }

        public string Name
        {
            get { return this.InvokePropertyGet<string>(0); }
        }
    }
}
