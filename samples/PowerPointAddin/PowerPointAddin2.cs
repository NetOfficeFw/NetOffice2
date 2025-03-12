using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using NetOffice.Office;
using NetOffice.Office.Core;
using NetOffice.PowerPoint;

namespace NetOfficeSamples
{
    [ComVisible(true)]
    [Guid("CD92E3AE-72B2-4E9F-A148-54D5F2B00238")]
    [ProgId("NetOfficeSamples.PowerPointAddin2")]
    public class PowerPointAddin2 : NetOffice.PowerPoint.PowerPointAddin
    {
        protected override void OnConnection(Application application, ext_ConnectMode connectMode, COMAddIn addin, ref Array custom)
        {
            var description = addin.Description;
            var name = application.Name;
        }
    }
}
