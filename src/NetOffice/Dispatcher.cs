using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace NetOffice
{
    // IID_IDispatch
    [Guid("00020400-0000-0000-C000-000000000046")]
    public class Dispatcher
    {
        const int S_OK = 0;
        const int LCID_US = 1033;
        const int IDispatch_Invoke_Offset = 6;

        internal delegate int InvokeMethod(IntPtr pDisp, int dispIdMember, ref Guid riid, uint lcid, ushort wFlags, ref DISPPARAMS pDispParams, out object pVarResult, ref EXCEPINFO pExcepInfo, out uint pArgErr);

        private IntPtr dispPtr;

        public Dispatcher(object instance)
        {
            this.dispPtr = Marshal.GetIDispatchForObject(instance);
        }

        protected unsafe T InvokePropertyGet<T>(int dispId)
        {
            var vtPtr = Marshal.ReadIntPtr(this.dispPtr);
            var invokePtr = Marshal.ReadIntPtr(vtPtr + IDispatch_Invoke_Offset * IntPtr.Size);
            var invoke = (InvokeMethod)Marshal.GetDelegateForFunctionPointer(invokePtr, typeof(InvokeMethod));

            var wFlags = INVOKEKIND.INVOKE_PROPERTYGET;
            var riid = Guid.Empty;

            var pDispParams = new DISPPARAMS();
            var pExcepInfo = new EXCEPINFO();
            uint pArgErr = 0;

            var result = new object();
            int hr = invoke(dispPtr, dispId, ref riid, LCID_US, (ushort)wFlags, ref pDispParams, out result, ref pExcepInfo, out pArgErr);

            if (hr != S_OK)
            {
                throw new COMException($"Failed to invoke member with dispId={dispId}.", hr);
            }

            return (T)result;
        }
    }
}
