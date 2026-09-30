using System;
using System.Collections.Generic;
using System.Reflection;
using ImageLink;
public class TestPicture { public int Type=13; }
public class TestPictureRange { public int Count=1; public object Item(int i){return new TestPicture();} }
public class TestSelection { public int Type=2; public bool HasChildShapeRange=false; public TestPictureRange ShapeRange=new TestPictureRange(); }
public class TestWindow { public TestSelection Selection=new TestSelection(); }
public class TestPresentations { public int Count=1; }
public class TestPresentation { public string Path=@"C:\test"; public string FullName=@"C:\test\test.pptx"; public int ReadOnly=0; }
public class TestApplication { public TestWindow ActiveWindow=new TestWindow(); public TestPresentations Presentations=new TestPresentations(); public TestPresentation ActivePresentation=new TestPresentation(); }
class SelectionTests {
 static int checks;
 static object Call(AddIn addin,string name,params object[] args) {
  try{return typeof(AddIn).GetMethod(name,BindingFlags.Instance|BindingFlags.NonPublic).Invoke(addin,args);}
  catch(TargetInvocationException e){throw e.InnerException;}
 }
 static void Check(bool value,string message){if(!value)throw new Exception(message);checks++;Console.WriteLine("PASS: "+message);}
 static void Reject(Action action,string expected){try{action();}catch(InvalidOperationException e){Check(e.Message.Contains(expected),"guard: "+expected);return;}throw new Exception("Expected guard: "+expected);}
 [STAThread] static int Main(){
  try{return Run();}catch(Exception e){Console.WriteLine("FAIL: "+e.ToString());return 1;}
 }
 static int Run(){
  var app=new TestApplication();var addin=new AddIn();Array custom=new object[0];addin.OnConnection(app,1,null,ref custom);
  Check(Object.ReferenceEquals(Call(addin,"Presentation"),app.ActivePresentation),"editable presentation accepted");
  Check(((List<object>)Call(addin,"Selection",false)).Count==1,"Boolean false allows normal picture (Create/Join/Detach)");
  Check(((List<object>)Call(addin,"Selection",true)).Count==1,"Boolean false allows single picture (Show/Update)");
  app.ActiveWindow.Selection.ShapeRange.Count=2;
  Check(((List<object>)Call(addin,"Selection",false)).Count==2,"multiple pictures accepted");
  Reject(()=>Call(addin,"Selection",true),"一张");
  app.ActiveWindow.Selection.HasChildShapeRange=true;
  Reject(()=>Call(addin,"Selection",false),"组合");
  app.ActiveWindow.Selection.HasChildShapeRange=false;app.ActiveWindow.Selection.Type=1;
  Reject(()=>Call(addin,"Selection",false),"选中图片");
  app.ActivePresentation.ReadOnly=-1;Reject(()=>Call(addin,"Presentation"),"只读");
  Console.WriteLine("ALL PASSED: "+checks+" selection/presentation guard checks");return 0;
 }
}
