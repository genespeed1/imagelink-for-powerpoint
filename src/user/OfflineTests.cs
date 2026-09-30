using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using ImageLink;

public class Tags
{
    Dictionary<string,string> values=new Dictionary<string,string>();
    public string Item(string key){return values.ContainsKey(key)?values[key]:"";}
    public void Add(string key,string value){values[key]=value;}
    public void Delete(string key){values.Remove(key);}
    public int Count {get{return values.Count;}}
    public string Name(int i){return values.Keys.ElementAt(i-1);}
    public string Value(int i){return values.Values.ElementAt(i-1);}
}
public class Crop
{
    public float ShapeWidth=240,ShapeHeight=180,PictureWidth=600,PictureHeight=300,PictureOffsetX=-90,PictureOffsetY=25;
}
public class PictureFormat {public Crop Crop=new Crop();}
public class Shape
{
    public float Left=30,Top=70,Width=240,Height=180,Rotation=23;
    public int HorizontalFlip=-1,VerticalFlip=0,Visible=-1,LockAspectRatio=-1,Type=13,Id=7,AutoShapeType=1;
    public string Name="test",AlternativeText="alt",Title="title";
    public Tags Tags=new Tags();public PictureFormat PictureFormat=new PictureFormat();
    public Shapes Parent;public Shapes GroupItems=new Shapes();public bool Deleted;
    public int AnimationCount;
    public int ZOrderPosition{get{return Parent==null?2:Parent.Items.IndexOf(this)+1;}}
    public void Flip(int axis){if(axis==0)HorizontalFlip=HorizontalFlip==0?-1:0;else VerticalFlip=VerticalFlip==0?-1:0;}
    public void PickUp(){}
    public void Apply(){}
    static int pickedAnimationCount;
    public void PickupAnimation(){if(AnimationCount==0)throw new InvalidOperationException("Invalid Request");pickedAnimationCount=AnimationCount;}
    public void ApplyAnimation(){AnimationCount=pickedAnimationCount;}
    public void ZOrder(int command){int i=Parent.Items.IndexOf(this);Parent.Items.RemoveAt(i);Parent.Items.Insert(i-1,this);}
    public void Delete(){Deleted=true;Parent.Items.Remove(this);}
}
public class Shapes
{
    public List<Shape> Items=new List<Shape>();public bool FailAdd;
    public int Count {get{return Items.Count;}}
    public Shape Item(int i){return Items[i-1];}
    public Shape AddPicture(string file,int link,int save,float left,float top,float w,float h)
    {
        if(FailAdd)throw new IOException("Injected insert failure");
        var s=new Shape{Parent=this,Width=1200,Height=600,HorizontalFlip=0,VerticalFlip=0};Items.Add(s);return s;
    }
}
public class Effect{public Shape Shape;}
public class Sequence
{
    public List<Effect> Items=new List<Effect>();
    public int Count{get{return Items.Count;}}
    public Effect Item(int i){return Items[i-1];}
}
public class Sequences
{
    public List<Sequence> Items=new List<Sequence>();
    public int Count{get{return Items.Count;}}
    public Sequence Item(int i){return Items[i-1];}
}
public class Timeline{public Sequence MainSequence=new Sequence();public Sequences InteractiveSequences=new Sequences();}
public class Slide{public int SlideIndex=1;public Shapes Shapes=new Shapes();public Timeline TimeLine=new Timeline();}
public class App{public void StartNewUndoEntry(){}}

static class OfflineTests
{
    static int count;
    static void Check(bool b,string name){if(!b)throw new Exception("FAIL "+name);count++;Console.WriteLine("PASS "+name);}
    static bool Near(double a,double b){return Math.Abs(a-b)<0.01;}
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetUi(IntPtr self,[MarshalAs(UnmanagedType.BStr)] string id,[MarshalAs(UnmanagedType.BStr)] out string xml);
    [STAThread] static int Main(string[] args)
    {
        try
        {
            var s=new Shape();var g=Geometry.Read(s);var fresh=new Shape{HorizontalFlip=0};
            g.Apply(fresh,2,false);
            Check(Near(fresh.PictureFormat.Crop.PictureWidth,600)&&Near(fresh.PictureFormat.Crop.PictureHeight,300),"same ratio keeps scaled picture dimensions");
            Check(Near(fresh.PictureFormat.Crop.PictureOffsetX,-90)&&Near(fresh.PictureFormat.Crop.PictureOffsetY,25),"same ratio keeps crop focus");
            Check(Near(fresh.Left,30)&&Near(fresh.Top,70)&&Near(fresh.Rotation,23)&&fresh.HorizontalFlip==-1,"frame rotation flip retained");
            g.Apply(fresh,2.0/3,false);var c=fresh.PictureFormat.Crop;
            Check(Near(c.PictureWidth/c.PictureHeight,2.0/3),"changed ratio stays proportional");
            Check(c.PictureWidth>=240&&c.PictureHeight>=180,"changed ratio covers frame");
            Check(Math.Abs(c.PictureOffsetX)<=(c.PictureWidth-240)/2+0.01&&Math.Abs(c.PictureOffsetY)<=(c.PictureHeight-180)/2+0.01,"offset constrained to image bounds");
            g.Apply(fresh,2,true);Check(Near(c.PictureOffsetX,0)&&Near(c.PictureOffsetY,0),"center fill resets crop offsets");
            Check(Near(c.PictureWidth,360)&&Near(c.PictureHeight,180),"center fill uses minimum covering dimensions");
            var circle=new Shape{AutoShapeType=9};var circleGeometry=Geometry.Read(circle);var circleReplacement=new Shape{AutoShapeType=1};
            circleGeometry.Apply(circleReplacement,2,false);Check(circleReplacement.AutoShapeType==9,"crop-to-shape type is preserved");
            var a=Target(1);var b=Target(2);Engine.Assign(a.Shape,"ASSET","组");Engine.Assign(b.Shape,"ASSET","组");
            b.Slide.Shapes.FailAdd=true;bool failed=false;
            try{Engine.Replace(new App(),new List<LinkedPicture>{a,b},"test.png",false);}catch{failed=true;}
            Check(failed&&!a.Shape.Deleted&&!b.Shape.Deleted,"stage failure keeps all originals");
            Check(a.Slide.Shapes.Count==1&&b.Slide.Shapes.Count==1,"stage failure removes all temporary replacements");
            var plain=Target(3);Engine.Replace(new App(),new List<LinkedPicture>{plain},"test.png",false);
            Check(plain.Shape.Deleted&&plain.Slide.Shapes.Count==1,"picture without animation skips PickupAnimation");
            b.Slide.Shapes.FailAdd=false;a.Shape.AnimationCount=2;b.Shape.AnimationCount=1;
            a.Slide.TimeLine.MainSequence.Items.Add(new Effect{Shape=a.Shape});
            b.Slide.TimeLine.MainSequence.Items.Add(new Effect{Shape=b.Shape});
            Engine.Replace(new App(),new List<LinkedPicture>{a,b},"test.png",false);
            Check(a.Shape.Deleted&&b.Shape.Deleted,"successful transaction removes old pictures");
            Check(a.Slide.Shapes.Count==1&&b.Slide.Shapes.Count==1,"successful replacement keeps shape counts");
            Check(Engine.Tag(a.Slide.Shapes.Item(1),Engine.IdTag)=="ASSET","association survives replacement");
            Check(a.Slide.Shapes.Item(1).Visible==-1,"replacement is visible");
            Check(a.Slide.Shapes.Item(1).AlternativeText=="alt","alternative text preserved");
            Check(a.Slide.Shapes.Item(1).AnimationCount==2&&b.Slide.Shapes.Item(1).AnimationCount==1,"animations copied to replacements");
            Engine.Detach(a.Slide.Shapes.Item(1));Check(Engine.Tag(a.Slide.Shapes.Item(1),Engine.IdTag)=="","detach clears binding");
            var target=Target(1);string fixture=Path.Combine(Path.GetFullPath(args[0]),"fixture.pptx");
            BuildPackage(fixture,"");Engine.Preflight(fixture,new List<LinkedPicture>{target});Check(true,"normal picture passes XML validation");
            BuildPackage(fixture,"<a:extLst><a:ext uri='test'><a14:useLocalDpi xmlns:a14='http://schemas.microsoft.com/office/drawing/2010/main' val='0'/></a:ext></a:extLst>");
            Engine.Preflight(fixture,new List<LinkedPicture>{target});Check(true,"normal Office local DPI extension accepted");
            BuildPackage(fixture,"<a:grayscl/>");CheckBlocked(fixture,target,"unsupported image effect rejected");
            BuildPackage(fixture,"", "ellipse");Engine.Preflight(fixture,new List<LinkedPicture>{target});Check(true,"preset crop shape accepted");
            BuildPackage(fixture,"", "custom");CheckBlocked(fixture,target,"custom freeform crop rejected");
            BuildPackage(fixture,"", "rect",true);Engine.Preflight(fixture,new List<LinkedPicture>{target});Check(true,"animation target accepted");
            BuildPackage(fixture,"");target.InGroup=true;CheckBlocked(fixture,target,"grouped picture rejected");
            var addin=new AddIn();var xml=XDocument.Parse(addin.GetCustomUI("Microsoft.PowerPoint.Presentation"));
            foreach(var action in xml.Descendants().Attributes("onAction"))Check(typeof(IImageLinkCallbacks).GetMethod(action.Value)!=null,"callback exists: "+action.Value);
            // Verify the actual dual COM vtable layout PowerPoint uses, without installing the add-in.
            IntPtr pointer=Marshal.GetComInterfaceForObject(addin,typeof(IRibbonExtensibility));
            try
            {
                IntPtr vtable=Marshal.ReadIntPtr(pointer);IntPtr function=Marshal.ReadIntPtr(vtable,7*IntPtr.Size);
                var callback=(GetUi)Marshal.GetDelegateForFunctionPointer(function,typeof(GetUi));string actual;
                int hr=callback(pointer,"Microsoft.PowerPoint.Presentation",out actual);
                Check(hr==0&&actual.Contains("imageLinkTab"),"native Ribbon COM vtable returns XML");
            }
            finally{Marshal.Release(pointer);}
            Console.WriteLine("ALL PASSED: "+count+" offline checks. PowerPoint integration NOT executed.");return 0;
        }
        catch(Exception ex){Console.WriteLine(ex);return 1;}
    }
    static LinkedPicture Target(int index)
    {
        var slide=new Slide{SlideIndex=index};var shape=new Shape{Parent=slide.Shapes};slide.Shapes.Items.Add(shape);
        return new LinkedPicture{Shape=shape,Slide=slide};
    }
    static void CheckBlocked(string file,LinkedPicture target,string name)
    {bool blocked=false;try{Engine.Preflight(file,new List<LinkedPicture>{target});}catch(InvalidOperationException){blocked=true;}Check(blocked,name);}
    static void BuildPackage(string file,string effects,string geometry="rect",bool animation=false)
    {
        if(File.Exists(file))File.Delete(file);
        using(var zip=ZipFile.Open(file,ZipArchiveMode.Create))
        {
            Entry(zip,"ppt/presentation.xml","<p:presentation xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><p:sldIdLst><p:sldId id='256' r:id='rId1'/></p:sldIdLst></p:presentation>");
            Entry(zip,"ppt/_rels/presentation.xml.rels","<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='slides/slide1.xml'/></Relationships>");
            string shapeGeometry=geometry=="custom"?"<a:custGeom/>":"<a:prstGeom prst='"+geometry+"'/>";
            Entry(zip,"ppt/slides/slide1.xml","<p:sld xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id='7' name='test'/><p:cNvPicPr><a:picLocks noChangeAspect='1'/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip>"+effects+"</a:blip></p:blipFill><p:spPr>"+shapeGeometry+"</p:spPr></p:pic></p:spTree></p:cSld>"+(animation?"<p:timing><p:spTgt spid='7'/></p:timing>":"")+"</p:sld>");
        }
    }
    static void Entry(ZipArchive zip,string path,string text){using(var writer=new StreamWriter(zip.CreateEntry(path).Open()))writer.Write(text);}
}

