#nullable enable

using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.Runtime;

namespace NetOffice.Office
{
    /// <summary>
    /// Object representing the Ribbon user interface (UI) of a Microsoft Office application.
    /// The object contains methods for invalidating control properties and for refreshing the user interface.
    /// </summary>
    /// <remarks>
    /// The reference to the object is returned by the onLoad procedure specified on the customUI tag.
    /// </remarks>
    [ComImport]
    [Guid(IID.IID_IRibbonUIGuid)]
    [TypeLibType(TypeLibTypeFlags.FDispatchable)]
    public interface IRibbonUI
    {
        /// <summary>
        /// Invalidates the cached values for all of the controls of the Ribbon user interface.
        /// </summary>
        [DispId(1)]
        void Invalidate();

        /// <summary>
        /// Invalidates the cached value for a single control on the Ribbon user interface.
        /// </summary>
        /// <param name="controlId">Specifies the ID of the control that will be invalidated</param>
        [DispId(2)]
        void InvalidateControl(string controlId);

        /// <summary>
        /// Used to invalidate a built-in control.
        /// </summary>
        /// <param name="controlId">Specified the identifier of the control that will be invalidated.</param>
        [DispId(3)]
        void InvalidateControlMso(string controlId);

        /// <summary>
        /// Activates the specified custom tab.
        /// </summary>
        /// <param name="controlId">Specifies the identifier of the custom Ribbon tab to be activated</param>
        [DispId(4)]
        void ActivateTab(string controlId);

        /// <summary>
        /// Activates the specified built-in tab.
        /// </summary>
        /// <param name="controlId">Specifies the identifier of the custom Ribbon tab to be activated.</param>
        [DispId(5)]
        void ActivateTabMso(string controlId);

        /// <summary>
        /// Activates the specified custom tab on the Microsoft Office Fluent Ribbon UI. Uses the fully qualified name of the tab which includes the identifier and the namespace of the tab.
        /// </summary>
        /// <param name="controlId">Specifies the identifier of the custom Ribbon tab to be activated</param>
        /// <param name="namespace">Specifies the namespace of the tab element</param>
        [DispId(6)]
        void ActivateTabQ(string controlId, string @namespace);
    }
}
