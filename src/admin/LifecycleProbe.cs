using System;
using System.Runtime.InteropServices;
using ImageLink;
class LifecycleProbe
{
 [DllImport("oleaut32.dll")] static extern IntPtr SafeArrayCreateVector(ushort vt,int lower,uint count);
 [DllImport("oleaut32.dll")] static extern int SafeArrayDestroy(IntPtr value);
 [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int Connect(IntPtr self,IntPtr application,int mode,IntPtr instance,ref IntPtr custom);
 [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int Event(IntPtr self,ref IntPtr custom);
 [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int Disconnect(IntPtr self,int mode,ref IntPtr custom);
 [STAThread] static int Main()
 {
  IntPtr ptr=Marshal.GetComInterfaceForObject(new AddIn(),typeof(IDTExtensibility2));
  IntPtr array=SafeArrayCreateVector(12,0,0);
  try
  {
   IntPtr table=Marshal.ReadIntPtr(ptr);
   var call=(Connect)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(table,7*IntPtr.Size),typeof(Connect));
   int result=call(ptr,IntPtr.Zero,1,IntPtr.Zero,ref array);
   Console.WriteLine("OnConnection HRESULT=0x"+result.ToString("X8"));
   if(result!=0) { Console.WriteLine(Marshal.GetExceptionForHR(result));return 1; }
   for(int i=9;i<=11;i++)
   {
    var evt=(Event)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(table,i*IntPtr.Size),typeof(Event));
    result=evt(ptr,ref array);Console.WriteLine("Lifecycle slot "+i+" HRESULT=0x"+result.ToString("X8"));
    if(result!=0)return 1;
   }
   var end=(Disconnect)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(table,8*IntPtr.Size),typeof(Disconnect));
   result=end(ptr,0,ref array);Console.WriteLine("OnDisconnection HRESULT=0x"+result.ToString("X8"));return result==0?0:1;
  }
  finally {SafeArrayDestroy(array);Marshal.Release(ptr);}
 }
}
