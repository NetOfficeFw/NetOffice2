namespace NetOffice.Runtime
{
    /// <summary>
    /// VT defines constants used in the <see cref="Variant._vt"/> discriminant field.
    /// </summary>
    /// <remarks>
    /// Values from the VARENUM enum from coreclr\pal\inc\rt\palrt.h header file.
    /// </remarks>
    public struct VT
    {
        public const int VT_EMPTY = 0;
        public const int VT_NULL = 1;
        public const int VT_I2 = 2;
        public const int VT_I4 = 3;
        public const int VT_R4 = 4;
        public const int VT_R8 = 5;
        public const int VT_CY = 6;

        /// <summary>
        /// OLE Automation date. Use <see cref="System.DateTime.ToOADate()"/> to set this value.
        /// </summary>
        /// <remarks>
        /// A variant time is stored as an 8-byte real value (double), representing a date between January 1, 100 and December 31, 9999, inclusive.
        /// The value 2.0 represents January 1, 1900; 3.0 represents January 2, 1900, and so on. Adding 1 to the value increments the date by a day.
        /// The fractional part of the value represents the time of day. Therefore, 2.5 represents noon on January 1, 1900;
        /// 3.25 represents 6:00 A.M. on January 2, 1900, and so on.
        /// Negative numbers represent the dates prior to December 30, 1899.
        /// </remarks>
        public const int VT_DATE = 7;
        
        /// <summary>
        /// BSTR is an OLE automation type for transferring length-prefixed strings, either Unicode or ANSI, as well as length-prefixed binary data.
        /// </summary>
        public const int VT_BSTR = 8;
        
        public const int VT_DISPATCH = 9;
        public const int VT_ERROR = 10;
        public const int VT_BOOL = 11;
        public const int VT_VARIANT = 12;
        public const int VT_UNKNOWN = 13;
        public const int VT_DECIMAL = 14;
        public const int VT_I1 = 16;
        public const int VT_UI1 = 17;
        public const int VT_UI2 = 18;
        public const int VT_UI4 = 19;
        public const int VT_I8 = 20;
        public const int VT_UI8 = 21;
        public const int VT_INT = 22;
        public const int VT_UINT = 23;
        public const int VT_VOID = 24;
        public const int VT_HRESULT = 25;
        public const int VT_PTR = 26;
        public const int VT_SAFEARRAY = 27;
        public const int VT_CARRAY = 28;
        public const int VT_USERDEFINED = 29;
        public const int VT_LPSTR = 30;
        public const int VT_LPWSTR = 31;
        public const int VT_RECORD = 36;
        public const int VT_INT_PTR = 37;
        public const int VT_UINT_PTR = 38;
    }
}
