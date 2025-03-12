using System;
using System.Runtime.InteropServices;

namespace NetOffice.Office.Core
{
    public class COMAddInClass : Dispatcher, COMAddIn
    {
        public COMAddInClass(object instance)
            : base(instance)
        {
        }

        public string Description
        {
            get { return this.InvokePropertyGet<string>(0); }
            set {; }
        }

        public string ProgId
        {
            get { return this.InvokePropertyGet<string>(3); }
        }

        public string Guid
        {
            get { return this.InvokePropertyGet<string>(4); }
        }

        public bool Connect
        {
            get { return this.InvokePropertyGet<bool>(6); }
            set {; }
        }

        public object Object
        {
            get { return this.InvokePropertyGet<object>(7); }
            set {; }
        }

        public object Parent
        {
            get { return this.InvokePropertyGet<object>(8); }
        }

        public object Application
        {
            get { return this.InvokePropertyGet<object>(1610743808); }
        }

        public int Creator
        {
            get { return this.InvokePropertyGet<int>(1610743809); }
        }
    }
}
