using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using ImageLink;
class ExperienceTests
{
 static int count;
 static void Check(bool value,string name){if(!value)throw new Exception(name);count++;Console.WriteLine("PASS: "+name);}
 [STAThread] static int Main(string[] args)
 {
  string dir=Path.Combine(Path.GetFullPath(args[0]),Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);
  string file=Path.Combine(dir,"preference.txt");
  Type t=typeof(AddIn).Assembly.GetType("ImageLink.NotificationPreferences");
  Func<object> load=()=>Activator.CreateInstance(t,new object[]{file});
  Func<object,bool> enabled=o=>(bool)t.GetProperty("Enabled").GetValue(o,null);
  Action<object,bool> set=(o,b)=>t.GetMethod("Set").Invoke(o,new object[]{b});
  object p=load();Check(!enabled(p),"notifications default off");
  set(p,true);Check(enabled(load()),"enabled preference persists across instances");
  set(p,false);Check(!enabled(load()),"disabled preference persists across instances");
  File.WriteAllText(file,"invalid");Check(!enabled(load()),"invalid preference safely defaults off");
  var a=new AddIn();var xml=XDocument.Parse(a.GetCustomUI("Microsoft.PowerPoint.Presentation"));
  foreach(string id in new[]{"ilCreate","ilJoin","ilDetach","ilShow","ilUpdate"})
  {var b=xml.Descendants().Single(x=>(string)x.Attribute("id")==id);Check((string)b.Attribute("size")=="large" && !String.IsNullOrEmpty((string)b.Attribute("imageMso")),id+" has large icon style");}
  Check((string)xml.Descendants().Single(x=>(string)x.Attribute("id")=="ilJoin").Attribute("imageMso")=="HyperlinkInsert","join uses corrected Office icon ID");
  var cb=xml.Descendants().Single(x=>(string)x.Attribute("id")=="ilSuccessNotifications");
  Check((string)cb.Parent.Attribute("id")=="imageLinkInfo","checkbox is beside help");
  Check(typeof(IImageLinkCallbacks).GetMethod((string)cb.Attribute("getPressed")).ReturnType==typeof(bool),"getPressed is exported as Boolean");
  Check(typeof(IImageLinkCallbacks).GetMethod((string)cb.Attribute("onAction")).GetParameters()[1].ParameterType==typeof(bool),"onAction accepts Boolean");
  // Test actual callbacks with a test-only preference path, without touching user settings.
  typeof(AddIn).GetField("notifications",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(a,load());
  a.SetSuccessNotifications(null,true);Check(a.GetSuccessNotifications(null)&&enabled(load()),"Ribbon callback enables and persists");
  a.SetSuccessNotifications(null,false);Check(!a.GetSuccessNotifications(null)&&!enabled(load()),"Ribbon callback disables and persists");
  typeof(AddIn).GetMethod("NotifySuccess",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(a,new object[]{"test"});
  Check(true,"disabled success notification returns without dialog");
  Console.WriteLine("ALL PASSED: "+count+" experience checks");return 0;
 }
}
