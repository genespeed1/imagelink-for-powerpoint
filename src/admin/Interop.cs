using System;
using System.Runtime.InteropServices;
using System.Reflection;

[assembly: AssemblyTitle("ImageLink for PowerPoint")]
[assembly: AssemblyDescription("Linked image updates for Windows PowerPoint")]
[assembly: AssemblyVersion("0.2.0.0")]
[assembly: ComVisible(false)]

namespace ImageLink
{
    [ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int connectMode, [MarshalAs(UnmanagedType.IDispatch)] object addInInst, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection(int removeMode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }
    [ComVisible(true), Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] string GetCustomUI(string ribbonID);
    }
    [ComVisible(true), Guid("7E849F02-9BA7-4420-9885-1D7531E0E5B0"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IImageLinkCallbacks
    {
        [DispId(10)] void CreateGroup([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(11)] void JoinGroup([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(12)] void UpdateGroup([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(13)] void Detach([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(14)] void ShowGroup([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(15)] void ShowHelp([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(19)] void OnRibbonLoad([MarshalAs(UnmanagedType.IDispatch)] object ribbon);
        [DispId(17)] [return: MarshalAs(UnmanagedType.VariantBool)] bool GetSuccessNotifications([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(18)] void SetSuccessNotifications([MarshalAs(UnmanagedType.IDispatch)] object control, [MarshalAs(UnmanagedType.VariantBool)] bool pressed);
    }
}




