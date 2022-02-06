using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Runtime
{
    /// <summary>
    /// Reserved COM dispatch identifiers (DISPID) when calling IDispatcher.Invoke() method.
    /// </summary>
    /// <remarks>
    /// Source: um/OAIdl.h file from Windows 10 SDK.
    /// </remarks>
    public static class DISPID
    {
        /// <summary>
        /// Value reserved to indicate an "unknown" name.
        /// </summary>
        public const int DISPID_UNKNOWN = -1;

        /// <summary>
        /// Value reserved for the "value" property.
        /// </summary>
        public const int DISPID_VALUE = 0;

        /// <summary>
        /// Value reserved to indicate the param that is the right-hand-side (or "put" value) of a PropertyPut.
        /// </summary>
        public const int DISPID_PROPERTYPUT = -3;

        /// <summary>
        /// Value reserved for the standard "NewEnum" method.
        /// </summary>
        public const int DISPID_NEWENUM = -4;

        /// <summary>
        /// Value reserved for the standard "Evaluate" method.
        /// </summary>
        public const int DISPID_EVALUATE = -5;

    }
}
