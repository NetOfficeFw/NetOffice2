using System;
using System.Runtime.InteropServices;

namespace NetOffice.PowerPoint
{
    //[TypeLibType(4288)]
    [Guid("91493442-5a91-11cf-8700-00aa0060263b")]
    public class _Application : Dispatcher
    {
        public _Application(object instance) : base(instance)
        {
        }

        public string Name
        {
            get { return this.InvokePropertyGet<string>(0); }
        }

        public string Caption
        {
            get { return this.InvokePropertyGet<string>(2009); }
            //set { return this.InvokePropertySet<string>(2009, value); }
        }
    }
}
