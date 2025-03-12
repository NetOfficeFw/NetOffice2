using System;

namespace NetOffice.Runtime
{
    /// <summary>
    /// Interface identifiers used by the OLE Automation protocols.
    /// </summary>
    /// <seealso cref="https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/58504586-e4af-44a3-be04-f1dc281b7429" />
    public static class IID
    {
        public const string IID_IUnknownGuid = "00000000-0000-0000-C000-000000000046";
        public const string IID_IDispatchGuid = "00020400-0000-0000-C000-000000000046";
        public const string IID_IDTExtensibility2Guid = "B65AD801-ABAF-11D0-BB8B-00A0C90F2744";
        public const string IID_IRibbonExtensibilityGuid = "000C0396-0000-0000-C000-000000000046";

        public static readonly Guid IID_NULL = Guid.Empty;
        public static readonly Guid IID_IUnknown = new Guid(IID_IUnknownGuid);
        public static readonly Guid IID_IDispatch = new Guid(IID_IUnknownGuid);
        public static readonly Guid IID_IDTExtensibility2 = new Guid(IID_IDTExtensibility2Guid);
    }
}
