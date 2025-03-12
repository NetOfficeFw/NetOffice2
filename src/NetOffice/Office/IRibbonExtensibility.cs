#nullable enable

using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.Runtime;

namespace NetOffice.Office
{
    /// <summary>
    /// The interface through which the Ribbon user interface (UI) communicates with a COM add-in to customize the UI.
    /// </summary>
    [ComImport]
    [Guid(IID.IID_IRibbonExtensibilityGuid)]
    [TypeLibType(TypeLibTypeFlags.FDispatchable)]
    public interface IRibbonExtensibility
    {
        /// <summary>
        /// Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.
        /// </summary>
        /// <param name="ribbonId">The identifier of the ribbon control.</param>
        [DispId(1)]
        string? GetCustomUI(string? ribbonId);
    }
}
