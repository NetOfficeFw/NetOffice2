using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using NetOffice.Office;
using NetOffice.Office.Core;
using NetOffice.PowerPoint;

namespace NetOfficeSamples
{
    [ComVisible(true)]
    [Guid("D8494D38-D995-4670-AFF9-9425ED71D657")]
    [ProgId("NetOfficeSamples.PowerPointAddin")]
    public class PowerPointAddin : IDTExtensibility2, IRibbonExtensibility
    {
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Trace.WriteLine($"Addin connected to application. Mode: {connectMode}");
            var type = application.GetType();
            var name = type.FullName;
            var isCom = Marshal.IsComObject(application);

            try
            {
                var addin = new COMAddInClass(addInInst);
                var description = addin.Description;
                // var p = addin.Parent;

                var ppApp = new Application(application);
                var appName = ppApp.Name;
                var caption = ppApp.Caption;

                var isVisible = ppApp.Visible;
                ppApp.Visible = MsoTriState.msoTrue;

                // ppApp.Caption = "Testing...";

                const int PP_CREATOR = 0x50575054;
                var creatorId = ppApp.Creator;

                if (PP_CREATOR == creatorId)
                {
                    Trace.WriteLine("It works!");
                }

                var height = ppApp.Height;
                var width = ppApp.Width;
                var sandbox = ppApp.IsSandboxed;
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

        public string GetCustomUI(string ribbonId)
        {
            var ribbon = /*lang=xml*/"""
                <?xml version="1.0" encoding="utf-8" ?>
                <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
                  <ribbon>
                    <tabs>
                      <!-- extend the Home tab -->
                      <tab idMso="TabHome">
                        <group id="MyGroupSample" label="Acme Group" insertAfterMso="GroupDrawing">
                          <button id="btnAction1"
                                  label="Action"
                                  size="large"
                                  imageMso="AppointmentColor5"
                                  />
                        </group>
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>
                """;

            return ribbon;
        }
    }
}
