using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using ImageLink;

static class SmokeTest
{
    static int assertions;
    static void Check(bool yes,string label) { if(!yes)throw new Exception("FAIL: "+label); assertions++;Console.WriteLine("PASS: "+label); }
    static void CloseEnough(double a,double b,string label) { Check(Math.Abs(a-b)<0.12,label+" ("+a+", "+b+")"); }
    static void Fixture(string file,int width,int height,Color color)
    {
        using(var b=new Bitmap(width,height))using(var g=Graphics.FromImage(b))
        {
            g.Clear(Color.White);g.FillRectangle(new SolidBrush(color),width/10,height/4,width*8/10,height/2);
            g.FillRectangle(Brushes.Black,width/5,height/2,width/10,height/4);
            g.FillEllipse(Brushes.Orange,width*6/10,height/3,width/7,height/3);
            for(int x=0;x<width;x+=width/10)g.DrawLine(Pens.LightGray,x,0,x,height);
            g.DrawString(Path.GetFileName(file),new Font("Arial",24),Brushes.Black,12,12);
            b.Save(file,ImageFormat.Png);
        }
    }
    [STAThread] public static int Main(string[] args)
    {
        string dir=Path.GetFullPath(args.Length>0?args[0]:Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"..","test"));Directory.CreateDirectory(dir);
        var log=new StreamWriter(Path.Combine(dir,"result.log"),false) { AutoFlush=true }; Console.SetOut(log);
        dynamic app=null,p=null;
        try
        {
            if(System.Diagnostics.Process.GetProcessesByName("POWERPNT").Length>0)
                throw new InvalidOperationException("Close PowerPoint before running integration tests. Existing user sessions will not be touched.");
            Fixture(Path.Combine(dir,"original.png"),1200,600,Color.SteelBlue);
            Fixture(Path.Combine(dir,"replacement.png"),2400,1200,Color.IndianRed);
            Fixture(Path.Combine(dir,"portrait.png"),600,900,Color.SeaGreen);
            app=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
            app.COMAddIns.Update();
            dynamic loadedAddin=app.COMAddIns.Item("ImageLink.PowerPointAddIn");
            loadedAddin.Connect=true;
            Check((bool)loadedAddin.Connect,"PowerPoint loads registered add-in");
            p=app.Presentations.Add(0);
            p.PageSetup.SlideWidth=960f;p.PageSetup.SlideHeight=540f;
            var originals=new List<Geometry>();
            for(int i=1;i<=4;i++)
            {
                dynamic slide=p.Slides.Add(i,12);
                slide.Shapes.AddShape(1,15f,15f,20f,20f);
                dynamic pic=slide.Shapes.AddPicture(Path.Combine(dir,"original.png"),0,-1,100f,100f,600f,300f);
                pic.Name="Linked_"+i;
                pic.Tags.Add("OTHER_TAG","PRESERVE");
                pic.AlternativeText="Test image "+i;pic.Title="Image title "+i;
                if(i>1)
                {
                    dynamic crop=pic.PictureFormat.Crop;
                    crop.ShapeWidth=i==2?240f:300f;crop.ShapeHeight=i==2?240f:180f;
                    crop.PictureWidth=i==2?600f:900f;crop.PictureHeight=i==2?300f:450f;
                    crop.PictureOffsetX=i==2?-70f:150f;crop.PictureOffsetY=i==2?10f:-60f;
                }
                if(i==3) {pic.Rotation=23f;pic.Flip(0);pic.Flip(1);}
                pic.Line.Visible=-1;pic.Line.ForeColor.RGB=0x0088ff;pic.Line.Weight=2.5f;
                if(i==2) {pic.Shadow.Visible=-1;pic.Shadow.OffsetX=3f;pic.Shadow.OffsetY=4f;}
                if(i<=3) Engine.Assign(pic,"TEST-ASSET","测试方案");
                else Engine.Assign(pic,"FROZEN-ASSET","历史版本");
                slide.Shapes.AddShape(1,760f,30f,25f,25f);
                originals.Add(Geometry.Read(pic));
            }
            string original=Path.Combine(dir,"original.pptx");p.SaveAs(original,24,0);
            List<LinkedPicture> targets=Engine.Find(p,"TEST-ASSET");Check(targets.Count==3,"find all three across slides");
            string backup=Path.Combine(dir,"test-snapshot.pptx");p.SaveCopyAs(backup,24,0);
            Engine.Preflight(backup,targets);Check(true,"preflight accepts normal crops / flips / border / shadow");
            int replaced=Engine.Replace(app,targets,Path.Combine(dir,"replacement.png"),false);Check(replaced==3,"replace count");
            targets=Engine.Find(p,"TEST-ASSET");
            for(int i=0;i<targets.Count;i++)
            {
                var before=originals[i];var after=Geometry.Read(targets[i].Shape);
                CloseEnough(before.Left,after.Left,"left "+i);CloseEnough(before.Top,after.Top,"top "+i);
                CloseEnough(before.Width,after.Width,"width "+i);CloseEnough(before.Height,after.Height,"height "+i);
                CloseEnough(before.PictureWidth,after.PictureWidth,"picture width "+i);CloseEnough(before.PictureHeight,after.PictureHeight,"picture height "+i);
                CloseEnough(before.OffsetX,after.OffsetX,"crop x "+i);CloseEnough(before.OffsetY,after.OffsetY,"crop y "+i);
                CloseEnough(before.Rotation,after.Rotation,"rotation "+i);
                Check(before.FlipH==after.FlipH&&before.FlipV==after.FlipV,"flips "+i);
                Check(before.ZOrder==after.ZOrder,"layer "+i);
                Check(Engine.Tag(targets[i].Shape,"OTHER_TAG")=="PRESERVE","other tags "+i);
                Check((string)targets[i].Shape.AlternativeText=="Test image "+(i+1),"alt text "+i);
                CloseEnough(2.5,(float)targets[i].Shape.Line.Weight,"line weight "+i);
                Check((int)targets[i].Shape.Line.ForeColor.RGB==0x0088ff,"line color "+i);
                Check((int)targets[i].Slide.Shapes.Count==3,"shape count no leaks "+i);
                targets[i].Slide.Export(Path.Combine(dir,"after_"+(i+1)+".png"),"PNG",1280,720);
            }
            Check((int)targets[1].Shape.Shadow.Visible==-1,"shadow remains visible");
            string updated=Path.Combine(dir,"updated.pptx");p.SaveAs(updated,24,0);
            CheckMedia(updated,Path.Combine(dir,"replacement.png"),3,Path.Combine(dir,"original.png"));
            Engine.Preflight(updated,targets);
            bool failed=false;
            try {Engine.Replace(app,targets,Path.Combine(dir,"missing.png"),false);}catch {failed=true;}
            Check(failed&&Engine.Find(p,"TEST-ASSET").Count==3,"invalid file does not delete originals");
            p.Close();p=null;
            p=app.Presentations.Open(updated,0,0,0);
            targets=Engine.Find(p,"TEST-ASSET");Check(targets.Count==3,"tags persist after reopen");
            dynamic duplicate=p.Slides.Item(1).Duplicate();
            Check(Engine.Find(p,"TEST-ASSET").Count==4,"duplicate slide retains association");
            duplicate.Item(1).Delete();
            targets=Engine.Find(p,"TEST-ASSET");
            Engine.Replace(app,targets,Path.Combine(dir,"portrait.png"),false);
            targets=Engine.Find(p,"TEST-ASSET");
            foreach(var t in targets)
            {
                var g=Geometry.Read(t.Shape);
                CloseEnough(g.PictureWidth/g.PictureHeight,600.0/900,"new aspect not stretched");
                Check(g.PictureWidth+0.1>=g.Width&&g.PictureHeight+0.1>=g.Height,"changed aspect covers frame");
            }
            p.SaveAs(Path.Combine(dir,"portrait.pptx"),24,0);
            targets[1].Slide.Export(Path.Combine(dir,"portrait_crop.png"),"PNG",1280,720);
            Engine.Replace(app,targets,Path.Combine(dir,"replacement.png"),true);
            targets=Engine.Find(p,"TEST-ASSET");
            foreach(var t in targets) {var g=Geometry.Read(t.Shape);CloseEnough(g.OffsetX,0,"center crop x");CloseEnough(g.OffsetY,0,"center crop y");}
            // Direct animations must transfer to the replacement picture.
            dynamic effect=targets[0].Slide.TimeLine.MainSequence.AddEffect(targets[0].Shape,1,0,1);
            string animated=Path.Combine(dir,"test-animated.pptx");p.SaveCopyAs(animated,24,0);failed=false;
            Engine.Preflight(animated,targets);Check(true,"animated target passes preflight");
            int effectCount=(int)targets[0].Slide.TimeLine.MainSequence.Count;
            Engine.Replace(app,targets,Path.Combine(dir,"portrait.png"),false);
            targets=Engine.Find(p,"TEST-ASSET");
            Check((int)targets[0].Slide.TimeLine.MainSequence.Count==effectCount,"animation count retained");
            Check((int)targets[0].Slide.TimeLine.MainSequence.Item(1).Shape.Id==(int)targets[0].Shape.Id,"animation targets replacement picture");
            effect=targets[0].Slide.TimeLine.MainSequence.Item(1);effect.Delete();
            Engine.Detach(targets[0].Shape);
            Check(Engine.Find(p,"TEST-ASSET").Count==2,"detach freezes one instance");
            Console.WriteLine("ALL PASSED: "+assertions+" assertions. PowerPoint "+app.Version);
            p.Saved=-1;p.Close();p=null;app.Quit();app=null;
            return 0;
        }
        catch(Exception ex) {Console.WriteLine(ex);return 1;}
        finally {if(p!=null){try{p.Saved=-1;p.Close();}catch{}}if(app!=null){try{app.Quit();}catch{}}}
    }
    static void CheckMedia(string file,string replacement,int count,string original)
    {
        XNamespace p="http://schemas.openxmlformats.org/presentationml/2006/main",a="http://schemas.openxmlformats.org/drawingml/2006/main",r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        using(var z=ZipFile.OpenRead(file))
        {
            for(int i=1;i<=4;i++)
            {
                XDocument s,rels;
                using(var stream=z.GetEntry("ppt/slides/slide"+i+".xml").Open())s=XDocument.Load(stream);
                using(var stream=z.GetEntry("ppt/slides/_rels/slide"+i+".xml.rels").Open())rels=XDocument.Load(stream);
                string rid=(string)s.Descendants(p+"pic").First().Descendants(a+"blip").First().Attribute(r+"embed");
                string target=(string)rels.Root.Elements().Single(x=>(string)x.Attribute("Id")==rid).Attribute("Target");
                string path=new Uri(new Uri("http://package/ppt/slides/slide"+i+".xml"),target).AbsolutePath.TrimStart('/');
                using(var memory=new MemoryStream())
                {using(var stream=z.GetEntry(path).Open())stream.CopyTo(memory);Check(memory.ToArray().SequenceEqual(File.ReadAllBytes(i<=count?replacement:original)),"actual media content slide "+i);}
            }
        }
    }
}
