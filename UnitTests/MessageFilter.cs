/// Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace Microsoft.VSPowerToys.ResourceRefactor.UnitTests
{
    /// <summary>
    /// Implementation of a custom COM message filter that resubmits calls
    /// failed due to application being busy.
    /// </summary>
    class MessageFilter : IOleMessageFilter
    {
        public static void Register()
        {
            IOleMessageFilter newfilter = new MessageFilter(); 

            IOleMessageFilter oldfilter = null; 
            CoRegisterMessageFilter(newfilter, out oldfilter);
        }

        public static void Revoke()
        {
            IOleMessageFilter oldfilter = null; 
            CoRegisterMessageFilter(null, out oldfilter);
        }

        //
        // IOleMessageFilter impl

        int IOleMessageFilter.HandleInComingCall(int dwCallType, System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr lpInterfaceInfo) 
        {
            return 0; //SERVERCALL_ISHANDLED
        }

        int IOleMessageFilter.RetryRejectedCall(System.IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == 2 ) //SERVERCALL_RETRYLATER
            {
                //Wait for application to become free
                System.Threading.Thread.Sleep(50);
                return 99; //retry immediately if return >=0 & <100
            }
            return -1; //cancel call
        }

        int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            System.Diagnostics.Debug.WriteLine("IOleMessageFilter::MessagePending");

            return 2; //PENDINGMSG_WAITDEFPROCESS 
        }

        //
        // Implementation

        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newfilter, out IOleMessageFilter oldfilter);
    }

    /// <summary>
    /// Decleration of COM MessageFilter interface
    /// </summary>
    [ComImport(), Guid("00000016-0000-0000-C000-000000000046"), 
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter // deliberately renamed to avoid confusion w/ System.Windows.Forms.IMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall( 
            int dwCallType, 
            IntPtr hTaskCaller, 
            int dwTickCount, 
            IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall( 
            IntPtr hTaskCallee, 
            int dwTickCount,
            int dwRejectType);

        [PreserveSig]
        int MessagePending( 
            IntPtr hTaskCallee, 
            int dwTickCount,
            int dwPendingType);
    }
}

