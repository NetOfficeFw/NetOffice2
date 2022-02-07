namespace NetOffice.Runtime
{
    /// <summary>
    /// VT defines constants used in the <see cref="Variant.vt"/> discriminant field.
    /// </summary>
    /// <remarks>
    /// Values from the VARENUM enum from coreclr\pal\inc\rt\palrt.h header file.
    /// </remarks>
    public struct VT
    {
        /// <summary>
        /// BSTR is an OLE automation type for transferring length-prefixed strings, either Unicode or ANSI, as well as length-prefixed binary data.
        /// </summary>
        public const int VT_BSTR = 8;
    }
}
