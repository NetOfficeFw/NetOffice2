using System;
using System.Runtime.InteropServices;

namespace NetOffice.Office
{
    [Guid("289E9AF2-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_DisconnectMode
    {
        /// <summary>
        /// The add-in was unloaded when Application was shut down.
        /// </summary>
        ext_dm_HostShutdown = 0,

        /// <summary>
        /// The add-in was unloaded while Application was running.
        /// </summary>
        ext_dm_UserClosed = 1,

        /// <summary>
        /// The add-in was unloaded after the user interface was set up.
        /// </summary>
        ext_dm_UISetupComplete = 2,

        /// <summary>
        /// The add-in was unloaded when the solution was closed.
        /// </summary>
        ext_dm_SolutionClosed = 3
    }
}
