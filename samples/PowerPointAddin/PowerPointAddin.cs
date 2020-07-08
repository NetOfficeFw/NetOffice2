using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using NetOffice.Office;

namespace NetOfficeSamples
{
    [ComVisible(true)]
    [Guid("D8494D38-D995-4670-AFF9-9425ED71D657")]
    [ProgId("NetOfficeSamples.PowerPointAddin")]
    public class PowerPointAddin : IDTExtensibility2
    {
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Trace.WriteLine($"Addin connected to application. Mode: {connectMode}");
            var type = application.GetType();
            var name = type.FullName;
            var isCom = Marshal.IsComObject(application);

            try
            {
                var unknown = Marshal.GetIUnknownForObject(application);
                //var dispatch = Marshal.GetIDispatchForObject(application);
            }
            catch (Exception ex)
            {
                Trace.TraceError($"Marshal failed. {ex}");
            }
        }

        public void OnDisconnection([In] ext_DisconnectMode removeMode, [In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
            Trace.WriteLine($"Addin disconnecting from application. Mode: {removeMode}");
        }

        public void OnAddInsUpdate([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }

        public void OnStartupComplete([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
            Trace.WriteLine($"Addin startup completed.");
        }

        public void OnBeginShutdown([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }
    }
}
