using System;
using System.Runtime.InteropServices;

namespace NetOffice.Word
{
    /// <summary>
    /// Microsoft Word application object.
    /// </summary>
    public class Application : _Application
    {
        public Application(object instance) : base(instance)
        {
        }
    }
}
