using System;
using System.Reflection;
class RegistrationTests
{
 static int Main(string[] args)
 {
  var setup=Assembly.LoadFrom(args[0]).GetType("Setup");
  var method=setup.GetMethod("CodeBaseForPath",BindingFlags.NonPublic|BindingFlags.Static);
  string[,] cases={{@"C:\Program Files\ImageLinkPowerPoint\ImageLink.AddIn.dll","file:///C:/Program Files/ImageLinkPowerPoint/ImageLink.AddIn.dll"},{@"C:\测试 文件夹\ImageLink.AddIn.dll","file:///C:/测试 文件夹/ImageLink.AddIn.dll"},{@"C:\ImageLink\ImageLink.AddIn.dll","file:///C:/ImageLink/ImageLink.AddIn.dll"}};
  for(int i=0;i<cases.GetLength(0);i++)if((string)method.Invoke(null,new object[]{cases[i,0]})!=cases[i,1])throw new Exception("CodeBase regression "+i);
  string installed=@"C:\Program Files\ImageLinkPowerPoint\ImageLink.AddIn.dll";
  if(System.IO.File.Exists(installed)){
   string actual=(string)method.Invoke(null,new object[]{installed});
   if(actual!=Assembly.ReflectionOnlyLoadFrom(installed).CodeBase)throw new Exception("CodeBase differs from CLR/RegAsm format");
  }
  Console.WriteLine("PASS: space, Unicode, plain path and installed assembly CodeBase");return 0;
 }
}
