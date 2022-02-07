using System;
using System.Runtime.InteropServices;

namespace NetOffice.Runtime
{
    /// <summary>
    /// Variant is the basic COM type for late-binding. It can contain any other COM data type.
    /// This type definition precisely matches the unmanaged data layout so that the struct can be passed
    /// to and from COM calls.
    /// </summary>
    /// <remarks>
    /// Variant size is the size of 4 pointers (16 bytes) on a 32-bit processor, and 3 pointers (24 bytes) on a 64-bit processor.
    /// 
    /// This struct definition must be binary compatible with the tagVARIANT type
    /// from coreclr\pal\inc\rt\palrt.h header file.
    /// </remarks>
    [StructLayout(LayoutKind.Explicit)]
    public struct Variant
    {
        [FieldOffset(0)]
        public ushort _vt;
        
        [FieldOffset(2)]
        private ushort _wReserved1;
        
        [FieldOffset(4)]
        private ushort _wReserved2;
        
        [FieldOffset(6)]
        private ushort _wReserved3;

        [FieldOffset(8)]
        public sbyte _i1;

        [FieldOffset(8)]
        public short _i2;

        [FieldOffset(8)]
        public int _i4;

        [FieldOffset(8)]
        public long _i8;

        [FieldOffset(8)]
        public byte _ui1;

        [FieldOffset(8)]
        public ushort _ui2;

        [FieldOffset(8)]
        public uint _ui4;

        [FieldOffset(8)]
        public ulong _ui8;

        [FieldOffset(8)]
        public int _int;

        [FieldOffset(8)]
        public uint _uint;

        [FieldOffset(8)]
        public short _bool;

        [FieldOffset(8)]
        public int _error;

        [FieldOffset(8)]
        public float _r4;

        [FieldOffset(8)]
        public double _r8;

        [FieldOffset(8)]
        public long _cy;

        [FieldOffset(8)]
        public double _date;

        [FieldOffset(8)]
        public IntPtr _bstr;

        [FieldOffset(8)]
        public IntPtr _unknown;

        [FieldOffset(8)]
        public IntPtr _dispatch;

        [FieldOffset(8)]
        public IntPtr _pvarVal;

        [FieldOffset(8)]
        public IntPtr _byref;

        [FieldOffset(8)]
        public Record _record;

        [StructLayout(LayoutKind.Sequential)]
        public struct Record
        {
            public IntPtr _record;
            public IntPtr _recordInfo;
        }
    }
}
