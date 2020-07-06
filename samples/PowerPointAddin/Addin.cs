using System;
using System.Runtime.InteropServices;
using NetOffice.Office;

namespace PowerPointAddin
{
    public class Addin : IDTExtensibility2
    {
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
        }

        public void OnDisconnection([In] ext_DisconnectMode removeMode, [In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }

        public void OnAddInsUpdate([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }

        public void OnStartupComplete([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }

        public void OnBeginShutdown([In, MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom)
        {
        }
    }
}
