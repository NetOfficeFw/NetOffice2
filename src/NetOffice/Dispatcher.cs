using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice.Runtime;

namespace NetOffice
{
    [Guid(IID.IID_IDispatchGuid)]
    public class Dispatcher
    {
        const int S_OK = 0;
        const int LCID_US = 1033;
        const int IDispatch_Invoke_Opnum = 6;

        internal delegate int InvokeMethod(IntPtr pDisp, int dispIdMember, ref Guid riid, uint lcid, ushort wFlags, ref DISPPARAMS pDispParams, out object pVarResult, ref EXCEPINFO pExcepInfo, out uint pArgErr);

        private IntPtr dispPtr;

        public Dispatcher(object instance)
        {
            this.dispPtr = Marshal.GetIDispatchForObject(instance);
        }

        protected unsafe T InvokePropertyGet<T>(int dispId)
        {
            var vtPtr = Marshal.ReadIntPtr(this.dispPtr);
            var invokePtr = Marshal.ReadIntPtr(vtPtr + IDispatch_Invoke_Opnum * IntPtr.Size);
            var invoke = (InvokeMethod)Marshal.GetDelegateForFunctionPointer(invokePtr, typeof(InvokeMethod));

            var riid = IID.IID_NULL;
            var wFlags = INVOKEKIND.INVOKE_PROPERTYGET;

            var pDispParams = new DISPPARAMS();
            var pExcepInfo = new EXCEPINFO();
            uint pArgErr = 0;

            var result = new object();
            int hr = invoke(dispPtr, dispId, ref riid, LCID_US, (ushort)wFlags, ref pDispParams, out result, ref pExcepInfo, out pArgErr);

            if (hr != S_OK)
            {
                throw new COMException($"Failed to invoke property get member with dispId={dispId}.", hr);
            }

            return (T)result;
        }

        protected unsafe void InvokePropertySet<T>(int dispId, string value)
        {
            var vtPtr = Marshal.ReadIntPtr(this.dispPtr);
            var invokePtr = Marshal.ReadIntPtr(vtPtr + IDispatch_Invoke_Opnum * IntPtr.Size);
            var invoke = (InvokeMethod)Marshal.GetDelegateForFunctionPointer(invokePtr, typeof(InvokeMethod));

            var riid = IID.IID_NULL;
            var wFlags = INVOKEKIND.INVOKE_PROPERTYPUT;

            // When you use IDispatch::Invoke() with DISPATCH_PROPERTYPUT or DISPATCH_PROPERTYPUTREF,
            // you have to specially initialize the cNamedArgs and rgdispidNamedArgs elements of your DISPPARAMS structure with the following:
            // 
            // DISPID dispidNamed = DISPID_PROPERTYPUT;
            // dispparams.cNamedArgs = 1;
            // dispparams.rgdispidNamedArgs = &dispidNamed;

            var dispidNamed = stackalloc int[1] { DISPID.DISPID_PROPERTYPUT };

            var pDispParams = new DISPPARAMS();
            pDispParams.cNamedArgs = 1;
            pDispParams.rgdispidNamedArgs = (IntPtr)dispidNamed;

            var pExcepInfo = new EXCEPINFO();
            uint pArgErr = 0;

            var varSize = Marshal.SizeOf<Variant>();
            var varValue = new Variant { _vt = VT.VT_BSTR, _bstr = Marshal.StringToBSTR(value) };

            IntPtr mem = Marshal.AllocCoTaskMem(varSize);
            Marshal.StructureToPtr(varValue, mem, false);

            pDispParams.cArgs = 1;
            pDispParams.rgvarg = mem;

            var result = new object();
            int hr = invoke(dispPtr, dispId, ref riid, LCID_US, (ushort)wFlags, ref pDispParams, out result, ref pExcepInfo, out pArgErr);

            Marshal.FreeBSTR(varValue._bstr);
            Marshal.FreeCoTaskMem(mem);

            if (hr != S_OK)
            {
                throw new COMException($"Failed to invoke property set member with dispId={dispId}.", hr);
            }
        }
    }
}
