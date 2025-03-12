using System;
using System.Runtime.InteropServices;

namespace NetOffice.Office.Core
{
    [Guid("000C0300-0000-0000-C000-000000000046")]
    public interface _IMsoDispObj
    {
        [DispId(1610743808)]
        object Application { get; }

        [DispId(1610743809)]
        int Creator { get; }
    }
}
