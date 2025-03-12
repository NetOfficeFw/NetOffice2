using System;
using System.Runtime.InteropServices;

namespace NetOffice.Office.Core
{
    [Guid("000C033A-0000-0000-C000-000000000046")]
    public interface COMAddIn : _IMsoDispObj
    {
        [DispId(0)]
        string Description { get; set; }

        [DispId(3)]
        string ProgId { get; }

        [DispId(4)]
        string Guid { get; }

        [DispId(6)]
        bool Connect { get; set; }

        [DispId(7)]
        object Object { get; set; }

        [DispId(8)]
        object Parent { get; }
    }
}
