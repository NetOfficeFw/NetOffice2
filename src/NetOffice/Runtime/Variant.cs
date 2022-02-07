using System;
using System.Runtime.InteropServices;

namespace NetOffice.Runtime
{
    /// <summary>
    /// A Variant is a special data type that can contain any kind of data.
    /// </summary>
    /// <remarks>
    /// This struct definition must be binary compatible with the tagVARIANT type
    /// from coreclr\pal\inc\rt\palrt.h header file.
    /// </remarks>
    [StructLayout(LayoutKind.Explicit, Size = 24)]
    public struct Variant
    {
        [FieldOffset(0)]
        public ushort vt;

        [FieldOffset(8)]
        public IntPtr bstr;
    }
}
