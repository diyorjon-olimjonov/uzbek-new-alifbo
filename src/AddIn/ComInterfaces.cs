// Office COM interfeyslari. Office PIA (office.dll, extensibility.dll) ga bog'lanmaslik
// uchun shu yerda e'lon qilingan - shu tufayli plagin Word 2010 dan Microsoft 365 gacha
// bitta DLL bilan ishlaydi. GUID va DispId'lar Office tip kutubxonalaridagi bilan bir xil.
using System;
using System.Runtime.InteropServices;

namespace Alifbo.Word
{
    [ComVisible(true)]
    public enum ExtConnectMode
    {
        AfterStartup = 0,
        Startup = 1,
        External = 2,
        CommandLine = 3
    }

    [ComVisible(true)]
    public enum ExtDisconnectMode
    {
        HostShutdown = 0,
        UserClosed = 1
    }

    [ComVisible(true)]
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] ExtConnectMode connectMode,
                          [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst, [In] ref Array custom);

        [DispId(2)]
        void OnDisconnection([In] ExtDisconnectMode removeMode, [In] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate([In] ref Array custom);

        [DispId(4)]
        void OnStartupComplete([In] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown([In] ref Array custom);
    }

    [ComVisible(true)]
    [ComImport, Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }
}
