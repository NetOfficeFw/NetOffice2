using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using NetOffice.Office;
using NetOffice.Word;

namespace NetOfficeSamples
{
    [ComVisible(true)]
    [Guid("AACDADCA-3C28-4F0F-B155-35CAAF9ECF87")]
    [ProgId("NetOfficeSamples.WordAddin")]
    public class WordAddin : IDTExtensibility2
    {
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Trace.WriteLine($"Addin connected to application. Mode: {connectMode}");
            
            try
            {
                var ppApp = new Application(application);
                var appName = ppApp.Name;
            }
            catch (Exception ex)
            {
                Trace.TraceError($"Addin failed. {ex}");
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
