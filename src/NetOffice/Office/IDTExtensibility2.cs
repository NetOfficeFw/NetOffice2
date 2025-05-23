﻿using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.Runtime;

namespace NetOffice.Office
{
    /// <summary>
    /// IDTExtensibility2 contains methods that act as interface between Microsoft Office applications and the add-in.
    /// Microsoft Office applications call these methods whenever an event that affects an add-in occurs,
    /// such as when it is loaded or unloaded.
    /// </summary>
    [ComImport]
    [Guid(IID.IID_IDTExtensibility2Guid)]
    [TypeLibType(TypeLibTypeFlags.FDispatchable | TypeLibTypeFlags.FDual)]
    public interface IDTExtensibility2
    {
        /// <summary>
        /// Occurs whenever an add-in is loaded into Microsoft Office application.
        /// </summary>
        /// <param name="application">A reference to an instance of the office application</param>
        /// <param name="connectMode">An ext_ConnectMode enumeration value that indicates the way the add-in was loaded into MS-Office</param>
        /// <param name="addInInst">An AddIn reference to the add-in's own instance. This is stored for later use, such as determining the parent collection for the add-in</param>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        [DispId(1)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnConnection([MarshalAs(26)][In] object application, [In] ext_ConnectMode connectMode, [MarshalAs(26)][In] object addInInst, [MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        /// <summary>
        /// Occurs whenever an add-in is unloaded from Microsoft Office application.
        /// </summary>
        /// <param name="removeMode">An ext_DisconnectMode enumeration value that informs an add-in why it was unloaded.</param>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use after the add-in unloads</param>
        [DispId(2)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnDisconnection([In] ext_DisconnectMode removeMode, [MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        /// <summary>
        /// Occurs whenever an add-in is loaded or unloaded Microsoft Office.
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        [DispId(3)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnAddInsUpdate([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        /// <summary>
        ///  Occurs whenever an add-in, which is set to load when Microsoft Office application starts, loads.
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use when the add-in loads</param>
        [DispId(4)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnStartupComplete([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);

        /// <summary>
        /// Occurs whenever Microsoft Office application shuts down while an add-in is running.
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        [DispId(5)]
        [MethodImpl(MethodImplOptions.InternalCall)]
        void OnBeginShutdown([MarshalAs(29, SafeArraySubType = VarEnum.VT_VARIANT)][In] ref Array custom);
    }
}
