#nullable enable

using System;
using System.Runtime.InteropServices;
using NetOffice.Runtime;

namespace NetOffice.Office
{
    /// <summary>
    /// Represents the control on a Ribbon user user interface (UI).
    /// </summary>
    [ComImport]
    [Guid(IID.IID_IRibbonControlGuid)]
    [TypeLibType(TypeLibTypeFlags.FDispatchable)]
    public interface IRibbonControl
    {
        /// <summary>
        /// Gets the ID of the control specified in the Ribbon XML markup customization file.
        /// </summary>
        [DispId(1)]
        string Id { get; }

        /// <summary>
        /// Represents the active window containing the Ribbon user interface that triggers a callback procedure.
        /// </summary>
        [DispId(2)]
        object? Context { [return: MarshalAs(UnmanagedType.IDispatch)] get; }

        /// <summary>
        /// Used to store arbitrary strings and fetch them at runtime.
        /// </summary>
        [DispId(3)]
        string? Tag { get; }
    }
}
