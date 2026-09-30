using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using ImageLink;
public class SnapshotPresentation {
 public string LastPath;public bool Fail,Animation;
 public void SaveCopyAs(string file,int format,int fonts){LastPath=file;
  typeof(OfflineTests).GetMethod("BuildPackage",BindingFlags.Static|BindingFlags.NonPublic).Invoke(null,new object[]{file,"","rect",Animation});
  if(Fail)throw new IOException("test save failure");
 }
}
public class TestRibbon {public int Count;public void Invalidate(){Count++;}}
class ContextTests {
 static int count;
 static void Check(bool ok,string name){if(!ok)throw new Exception(name);count++;Console.WriteLine("PASS "+name);}
 [STAThread] static int Main(string[] args){
 var addin=new AddIn();var xml=XDocument.Parse(addin.GetCustomUI("Microsoft.PowerPoint.Presentation"));
 var ids=xml.Descendants().Attributes("id").Select(x=>x.Value).ToList();Check(ids.Distinct().Count()==ids.Count,"unique Ribbon and context IDs");
 var ribbonActions=xml.Descendants().Single(x=>x.Name.LocalName=="ribbon").Descendants().Attributes("onAction").Select(x=>x.Value).OrderBy(x=>x).ToArray();
 foreach(var menu in xml.Descendants().Where(x=>x.Name.LocalName=="contextMenu")){
 Check(menu.Elements().Single().Name.LocalName=="menu","one parent menu");
 Check(menu.Descendants().Attributes("onAction").Select(x=>x.Value).SequenceEqual(new[]{"UpdateGroup","CreateGroup","JoinGroup","Detach","ShowGroup"}),"context menu has five actions in requested order");}
 Check(!xml.ToString().Contains("OpenBackups"),"backup action removed");
 var target=(LinkedPicture)typeof(OfflineTests).GetMethod("Target",BindingFlags.Static|BindingFlags.NonPublic).Invoke(null,new object[]{1});
 var p=new SnapshotPresentation();Engine.ValidateCurrentPresentation(p,new List<LinkedPicture>{target});Check(!File.Exists(p.LastPath),"snapshot removed after successful validation");
 p.Animation=true;Engine.ValidateCurrentPresentation(p,new List<LinkedPicture>{target});
 Check(!File.Exists(p.LastPath),"snapshot removed after animated-image validation");
 p.Fail=true;bool failed=false;try{Engine.ValidateCurrentPresentation(p,new List<LinkedPicture>{target});}catch(IOException){failed=true;}
 Check(failed&&!File.Exists(p.LastPath),"partial snapshot removed after failed save");
 var prefs=typeof(AddIn).Assembly.GetType("ImageLink.NotificationPreferences");var pref=Activator.CreateInstance(prefs,new object[]{Path.Combine(Path.GetFullPath(args[0]),"context-settings.txt")});
 typeof(AddIn).GetField("notifications",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(addin,pref);
 var ui=new TestRibbon();addin.OnRibbonLoad(ui);addin.SetSuccessNotifications(null,true);addin.SetSuccessNotifications(null,false);
 Check(ui.Count==2&&!addin.GetSuccessNotifications(null),"toolbar and menu notification state invalidated together");
 Console.WriteLine("ALL PASSED "+count+" context/snapshot checks");return 0;
 }
}
