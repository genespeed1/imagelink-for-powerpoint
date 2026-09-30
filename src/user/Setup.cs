using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Security.Principal;
using System.Windows.Forms;
using Microsoft.Win32;

static class Setup
{
 const string ProgId="ImageLink.PowerPointUserAddIn";
 const string Clsid="{B79D4FE3-C8B2-4B86-91DD-8CE0BBD00414}";
 static readonly string[] ClassKeys={"Software\\Classes\\CLSID\\"+Clsid,"Software\\Classes\\"+ProgId,"Software\\Microsoft\\Office\\PowerPoint\\Addins\\"+ProgId};
 static string Folder {get{return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"ImageLinkPowerPointUser");}}
 static RegistryView[] Views {get{return Environment.Is64BitOperatingSystem?new[]{RegistryView.Registry32,RegistryView.Registry64}:new[]{RegistryView.Registry32};}}
 [STAThread] static void Main()
 {
  Application.EnableVisualStyles();
  using(var f=new Form{Text="ImageLink 图片关联 · 0.2.0 免管理员版安装",ClientSize=new Size(600,335),Font=new Font("Microsoft YaHei UI",10),FormBorderStyle=FormBorderStyle.FixedDialog,MaximizeBox=false,StartPosition=FormStartPosition.CenterScreen,AutoScaleMode=AutoScaleMode.Dpi})
  {
   f.Controls.Add(new Label{Text="PowerPoint 图片关联 · 免管理员版",Font=new Font("Microsoft YaHei UI",17,FontStyle.Bold),Bounds=new Rectangle(24,20,550,42)});
   f.Controls.Add(new Label{Text="0.2.0 · 形状裁剪与动画保留版 · 仅当前用户安装\r\n\r\n直接双击运行，无需管理员权限。\r\n安装到当前用户的 LocalAppData 目录。\r\n仅注册到当前账户，不修改系统安装版。\r\n需要公司允许加载此类 Office 插件。\r\n\r\n安装或卸载前，请保存并关闭 PowerPoint。",Bounds=new Rectangle(24,74,550,185)});
   var install=new Button{Text="安装 / 修复",Bounds=new Rectangle(24,282,170,34)};
   var uninstall=new Button{Text="卸载插件",Bounds=new Rectangle(214,282,170,34)};
   var close=new Button{Text="关闭",Bounds=new Rectangle(404,282,170,34)};
   install.Click+=(s,e)=>Run(f,Install);uninstall.Click+=(s,e)=>Run(f,Uninstall);close.Click+=(s,e)=>f.Close();
   f.Controls.Add(install);f.Controls.Add(uninstall);f.Controls.Add(close);Application.Run(f);
  }
 }
 static void Run(Form owner,Func<string> action)
 {
  try
  {
   if(new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))throw new InvalidOperationException("本版适用于普通用户权限。请直接双击运行，不要选择以管理员身份运行。如果当前账户关闭了 UAC 并以管理员身份运行，请使用原系统安装版。");
   if(Process.GetProcessesByName("POWERPNT").Length>0)throw new InvalidOperationException("请先保存并完全退出 PowerPoint，再重试。不会自动关闭你的文稿。");
   MessageBox.Show(owner,action(),"ImageLink",MessageBoxButtons.OK,MessageBoxIcon.Information);
  }
  catch(Exception e){MessageBox.Show(owner,e.Message,"安装 / 卸载未完成",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
 }
 static string Install()
 {
  var saved=new List<RegistryBackup>();
  foreach(var view in Views)saved.Add(new RegistryBackup(RegistryHive.CurrentUser,view,"Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\ImageLinkPowerPointUser"));
  foreach(var view in Views)foreach(var hive in new[]{RegistryHive.CurrentUser})foreach(var key in ClassKeys)saved.Add(new RegistryBackup(hive,view,key));
  Directory.CreateDirectory(Folder);
  var oldFiles=new Dictionary<string,byte[]>();
  string[] payloads=Environment.Is64BitOperatingSystem?new[]{"ImageLink.UserAddIn.dll","ImageLink-Verify-x86.exe","ImageLink-Verify-x64.exe"}:new[]{"ImageLink.UserAddIn.dll","ImageLink-Verify-x86.exe"};
  foreach(string name in new[]{"ImageLink.UserAddIn.dll","ImageLink-Verify-x86.exe","ImageLink-Verify-x64.exe","ImageLink-Setup.exe"}){string path=Path.Combine(Folder,name);oldFiles[path]=File.Exists(path)?File.ReadAllBytes(path):null;}
  try
  {
   foreach(string name in payloads)using(var input=Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
   {if(input==null)throw new IOException("安装包不完整："+name);using(var output=File.Create(Path.Combine(Folder,name)))input.CopyTo(output);}
   string dll=Path.Combine(Folder,"ImageLink.UserAddIn.dll");string assemblyName=AssemblyName.GetAssemblyName(dll).FullName;
   foreach(var view in Views)
   {
    using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))foreach(string key in ClassKeys)user.DeleteSubKeyTree(key,false);
    using(var userRoot=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
    {
     using(var cls=userRoot.CreateSubKey(ClassKeys[0]))
     {
      cls.SetValue("","ImageLink PowerPoint Add-in");
      using(var server=cls.CreateSubKey("InprocServer32"))
      {
       server.SetValue("","mscoree.dll");server.SetValue("ThreadingModel","Both");SetAssembly(server,dll,assemblyName);
       using(var version=server.CreateSubKey("0.2.0.0"))SetAssembly(version,dll,assemblyName);
      }
      using(var prog=cls.CreateSubKey("ProgId"))prog.SetValue("",ProgId);
      using(var cat=cls.CreateSubKey("Implemented Categories\\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}")){}
     }
     using(var prog=userRoot.CreateSubKey(ClassKeys[1])){prog.SetValue("","ImageLink.AddIn");using(var id=prog.CreateSubKey("CLSID"))id.SetValue("",Clsid);}
     using(var addin=userRoot.CreateSubKey(ClassKeys[2]))
     {addin.SetValue("FriendlyName","ImageLink 图片关联（免管理员版）");addin.SetValue("Description","图片跨页关联更新 · 0.2.0");addin.SetValue("LoadBehavior",3,RegistryValueKind.DWord);}
    }
   }
   foreach(string bit in Environment.Is64BitOperatingSystem?new[]{"x86","x64"}:new[]{"x86"})
   {
    using(var process=Process.Start(new ProcessStartInfo(Path.Combine(Folder,"ImageLink-Verify-"+bit+".exe")){UseShellExecute=false,CreateNoWindow=true,WindowStyle=ProcessWindowStyle.Hidden}))
    {
     if(!process.WaitForExit(20000)){process.Kill();throw new InvalidOperationException(bit+" 加载自检超时。");}
     if(process.ExitCode!=0)throw new InvalidOperationException(bit+" 加载自检失败。详情：%LOCALAPPDATA%\\ImageLinkPowerPointUser\\Logs");
    }
   }
  string self=Assembly.GetExecutingAssembly().Location,installed=Path.Combine(Folder,"ImageLink-Setup.exe");
  if(!String.Equals(Path.GetFullPath(self),installed,StringComparison.OrdinalIgnoreCase))File.Copy(self,installed,true);
  using(var userRoot=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,Environment.Is64BitOperatingSystem?RegistryView.Registry64:RegistryView.Registry32))
  using(var key=userRoot.CreateSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\ImageLinkPowerPointUser"))
  {key.SetValue("DisplayName","ImageLink 图片关联（仅当前用户）");key.SetValue("DisplayVersion","0.2.0");key.SetValue("UninstallString","\""+installed+"\"");key.SetValue("InstallLocation",Folder);key.SetValue("NoModify",1,RegistryValueKind.DWord);key.SetValue("NoRepair",1,RegistryValueKind.DWord);}

  }
  catch(Exception error)
  {
   var rollbackErrors=new List<string>();
   foreach(var item in oldFiles)try{if(item.Value==null){if(File.Exists(item.Key))File.Delete(item.Key);}else File.WriteAllBytes(item.Key,item.Value);}catch(Exception e){rollbackErrors.Add(e.Message);}
   foreach(var backup in saved)try{backup.Restore();}catch(Exception e){rollbackErrors.Add(e.Message);}
   throw new InvalidOperationException(error.Message+"\r\n"+(rollbackErrors.Count==0?"已恢复安装前的插件文件与注册信息。":"回滚部分失败，请保留提示："+String.Join("; ",rollbackErrors)));
  }
  return "安装完成，当前用户 COM 加载自检通过（"+(Environment.Is64BitOperatingSystem?"32 位 / 64 位":"32 位")+"）。\r\n\r\n请重新打开 PowerPoint，检查“图片关联”选项卡。\r\n自检日志：%LOCALAPPDATA%\\ImageLinkPowerPointUser\\Logs";
 }
 static void SetAssembly(RegistryKey key,string dll,string name)
 {key.SetValue("Class","ImageLink.AddIn");key.SetValue("Assembly",name);key.SetValue("RuntimeVersion","v4.0.30319");key.SetValue("CodeBase",CodeBaseForPath(dll));}
 // Match RegAsm /codebase. Native Office activation fails with HRESULT 0x80070002
 // for escaped spaces (%20), even when an already-running CLR can load the DLL.
 internal static string CodeBaseForPath(string dll)
 {return new Uri(Path.GetFullPath(dll)).GetComponents(UriComponents.AbsoluteUri,UriFormat.Unescaped);}
 static string Uninstall()
 {
  foreach(var view in Views)foreach(var hive in new[]{RegistryHive.CurrentUser})using(var root=RegistryKey.OpenBaseKey(hive,view))foreach(string key in ClassKeys)root.DeleteSubKeyTree(key,false);
  using(var userRoot=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,Environment.Is64BitOperatingSystem?RegistryView.Registry64:RegistryView.Registry32))userRoot.DeleteSubKeyTree("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\ImageLinkPowerPointUser",false);
  foreach(string name in new[]{"ImageLink.UserAddIn.dll","ImageLink-Verify-x86.exe","ImageLink-Verify-x64.exe"}){string path=Path.Combine(Folder,name);if(File.Exists(path))File.Delete(path);}
  return "免管理员版已从当前账户卸载。PPT、自动备份和其他版本均保留。";
 }
 sealed class RegistryBackup
 {
  readonly RegistryHive hive;readonly RegistryView view;readonly string path;readonly Tree tree;
  public RegistryBackup(RegistryHive h,RegistryView v,string p){hive=h;view=v;path=p;using(var root=RegistryKey.OpenBaseKey(h,v))using(var key=root.OpenSubKey(p))tree=key==null?null:new Tree(key);}
  public void Restore(){using(var root=RegistryKey.OpenBaseKey(hive,view)){root.DeleteSubKeyTree(path,false);if(tree!=null)using(var key=root.CreateSubKey(path))tree.Write(key);}}
 }
 sealed class Tree
 {
  readonly Dictionary<string,object> values=new Dictionary<string,object>();readonly Dictionary<string,RegistryValueKind> kinds=new Dictionary<string,RegistryValueKind>();readonly Dictionary<string,Tree> children=new Dictionary<string,Tree>();
  public Tree(RegistryKey key){foreach(string name in key.GetValueNames()){values[name]=key.GetValue(name,null,RegistryValueOptions.DoNotExpandEnvironmentNames);kinds[name]=key.GetValueKind(name);}foreach(string name in key.GetSubKeyNames())using(var child=key.OpenSubKey(name))children[name]=new Tree(child);}
  public void Write(RegistryKey key){foreach(var pair in values)key.SetValue(pair.Key,pair.Value,kinds[pair.Key]);foreach(var pair in children)using(var child=key.CreateSubKey(pair.Key))pair.Value.Write(child);}
 }
}





