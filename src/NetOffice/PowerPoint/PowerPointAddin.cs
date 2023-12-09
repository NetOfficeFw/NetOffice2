using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using NetOffice.Office;
using NetOffice.Office.Core;

namespace NetOffice.PowerPoint
{
    public class PowerPointAddin : IDTExtensibility2
    {
        
        [SupportedOSPlatform("windows")]
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                var ppApp = new Application(application);
                var addin = new COMAddInClass(addInInst);


                this.OnConnection(ppApp, connectMode, addin, ref custom);
            }
            catch (Exception ex)
            {
                Trace.TraceError($"Addin failed. {ex}");
            }
        }

        public void OnDisconnection([In] ext_DisconnectMode removeMode, [In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
            throw new NotImplementedException();
        }

        public void OnAddInsUpdate([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
            throw new NotImplementedException();
        }

        public void OnStartupComplete([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
            throw new NotImplementedException();
        }

        public void OnBeginShutdown([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
            throw new NotImplementedException();
        }

        protected virtual void OnConnection(Application application, ext_ConnectMode connectMode, COMAddIn addin, ref Array custom)
        {
        }
    }
}
