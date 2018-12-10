using Microsoft.Win32.SafeHandles;
using System;
using System.Runtime.InteropServices;

namespace VbsWriter
{
    public class workProcedure : IDisposable
    {
        public string startRow { get;  set; }
        public string no { get; set; }
        public string startTime { get; set; }
        public string endTime { get; set; }
        public string totalTime { get; set; }
        public string targetEnvironment { get; set; }
        public string targetArea { get; set; }
        public string workTitle { get; set; }
        public string responsible { get; set; }
        public string workPlace { get; set; }
        public string procedure { get; set; }
        public string sap { get; set; }
        public string eni { get; set; }
        public string hako { get; set; }
        public string asa { get; set; }
        public string hiro { get; set; }
        public string aki { get; set; }

        public workProcedure()
        {
        }

        // Flag: Has Dispose already been called?
        bool disposed = false;
        // Instantiate a SafeHandle instance.
        SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);

        // Public implementation of Dispose pattern callable by consumers.
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // Protected implementation of Dispose pattern.
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                handle.Dispose();
                // Free any other managed objects here.
                //
            }

            disposed = true;
        }
    }
}
