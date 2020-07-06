using System;
using System.Runtime.InteropServices;

namespace NetOffice.Office
{
    [Guid("289E9AF1-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_ConnectMode
    {
        /// <summary>
        /// The add-in was loaded after Application started.
        /// </summary>
        ext_cm_AfterStartup = 0,

        /// <summary>
        /// The add-in was loaded when Application started.
        /// </summary>
        ext_cm_Startup = 1,

        /// <summary>
        /// The add-in was loaded by an external client.
        /// </summary>
        ext_cm_External = 2,

        /// <summary>
        /// The add-in was loaded from the command line.
        /// </summary>
        ext_cm_CommandLine = 3,

        /// <summary>
        /// The add-in was loaded with a solution.
        /// </summary>
        ext_cm_Solution = 4,

        /// <summary>
        /// The add-in was loaded for user interface setup.
        /// </summary>
        ext_cm_UISetup = 5
    }
}
