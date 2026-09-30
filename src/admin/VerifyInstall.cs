using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Linq;
class VerifyInstall
{
 [DllImport("oleaut32.dll")] static extern IntPtr SafeArrayCreateVector(ushort type,int low,uint count);
 [DllImport("oleaut32.dll")] static extern int SafeArrayDestroy(IntPtr value);
 [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int Connect(IntPtr self,IntPtr app,int mode,IntPtr addin,ref IntPtr custom);
 [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetUI(IntPtr self,[MarshalAs(UnmanagedType.BStr)] string ribbon,[MarshalAs(UnmanagedType.BStr)] out string xml);
 [STAThread] static int Main()
 {
  string folder=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"ImageLinkPowerPoint","Logs");Directory.CreateDirectory(folder);
  using(var log=new StreamWriter(Path.Combine(folder,"verify-"+(Environment.Is64BitProcess?"x64":"x86")+".txt"),false))
  {
   log.AutoFlush=true;
   try
   {
    log.WriteLine(DateTime.Now.ToString("s")+" CLR "+Environment.Version);
    using(var root=Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine,Environment.Is64BitProcess?Microsoft.Win32.RegistryView.Registry64:Microsoft.Win32.RegistryView.Registry32))
    using(var key=root.OpenSubKey("Software\\Classes\\CLSID\\{90622479-79E5-4C09-A4EF-87A136D3221A}\\InprocServer32"))
    {
     string codebase=key==null?null:key.GetValue("CodeBase") as string;
     if(String.IsNullOrEmpty(codebase))throw new Exception("Missing CodeBase");
     var uri=new Uri(codebase);
     if(codebase!=uri.GetComponents(UriComponents.AbsoluteUri,UriFormat.Unescaped))throw new Exception("CodeBase must match RegAsm's unescaped path for native Office activation.");
     if(!File.Exists(uri.LocalPath))throw new FileNotFoundException("Registered DLL not found",uri.LocalPath);
     log.WriteLine("Native activation path format OK");
    }
    object addin=Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("90622479-79E5-4C09-A4EF-87A136D3221A"),true));
    log.WriteLine("Registered activation OK: "+addin.GetType().Assembly.FullName);
    if(addin.GetType().Assembly.GetName().Version.ToString()!="0.2.0.0")throw new Exception("Unexpected installed version");
    IntPtr unknown=Marshal.GetIUnknownForObject(addin);
    try
    {
     Guid iid=new Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744");IntPtr lifecycle;
     Marshal.ThrowExceptionForHR(Marshal.QueryInterface(unknown,ref iid,out lifecycle));
     IntPtr array=SafeArrayCreateVector(12,0,0);
     try
     {
      var callback=(Connect)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(Marshal.ReadIntPtr(lifecycle),7*IntPtr.Size),typeof(Connect));
      Marshal.ThrowExceptionForHR(callback(lifecycle,IntPtr.Zero,1,IntPtr.Zero,ref array));log.WriteLine("Native OnConnection OK");
     }
     finally{SafeArrayDestroy(array);Marshal.Release(lifecycle);}
     iid=new Guid("000C0396-0000-0000-C000-000000000046");IntPtr ribbon;
     Marshal.ThrowExceptionForHR(Marshal.QueryInterface(unknown,ref iid,out ribbon));
     try
     {
      var callback=(GetUI)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(Marshal.ReadIntPtr(ribbon),7*IntPtr.Size),typeof(GetUI));
      string xml;Marshal.ThrowExceptionForHR(callback(ribbon,"Microsoft.PowerPoint.Presentation",out xml));XDocument.Parse(xml);log.WriteLine("Native Ribbon XML OK");
     }
     finally{Marshal.Release(ribbon);}
    }
    finally{Marshal.Release(unknown);}
    log.WriteLine("PASS");return 0;
   }
   catch(Exception e){log.WriteLine(e.ToString());return 1;}
  }
 }
}




