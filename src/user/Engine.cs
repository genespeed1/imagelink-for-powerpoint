using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace ImageLink
{
    public sealed class LinkedPicture
    {
        public dynamic Shape;
        public dynamic Slide;
        public bool InGroup;
        public int SlideIndex { get { return (int)Slide.SlideIndex; } }
        public string Label { get { return "第 " + SlideIndex + " 页 · " + (string)Shape.Name; } }
    }

    public sealed class Geometry
    {
        public float Left, Top, Width, Height, Rotation, PictureWidth, PictureHeight, OffsetX, OffsetY;
        public int FlipH, FlipV, Visibility, ZOrder, LockAspect, AutoShapeType;
        public static Geometry Read(dynamic s)
        {
            dynamic c = s.PictureFormat.Crop;
            int autoShapeType=-2;
            try { autoShapeType=(int)s.AutoShapeType; } catch { }
            return new Geometry { Left=(float)s.Left, Top=(float)s.Top, Width=(float)s.Width, Height=(float)s.Height,
                Rotation=(float)s.Rotation, PictureWidth=(float)c.PictureWidth, PictureHeight=(float)c.PictureHeight,
                OffsetX=(float)c.PictureOffsetX, OffsetY=(float)c.PictureOffsetY,
                FlipH=(int)s.HorizontalFlip, FlipV=(int)s.VerticalFlip, Visibility=(int)s.Visible,
                ZOrder=(int)s.ZOrderPosition, LockAspect=(int)s.LockAspectRatio, AutoShapeType=autoShapeType };
        }
        public void Apply(dynamic s, double newAspect, bool fill)
        {
            s.LockAspectRatio = 0;
            s.Rotation = 0f;
            s.Width = Width;
            s.Height = Height;
            // Crop-to-shape is stored as the picture's AutoShapeType. Restore it
            // before applying the crop box so circles and other preset masks survive.
            if(AutoShapeType!=-2) s.AutoShapeType=AutoShapeType;
            dynamic c = s.PictureFormat.Crop;
            c.ShapeWidth = Width;
            c.ShapeHeight = Height;
            double pw = PictureWidth, ph = PictureHeight, ox = OffsetX, oy = OffsetY;
            double oldAspect = PictureWidth / Math.Max(0.001, PictureHeight);
            if (fill || Math.Abs(oldAspect / newAspect - 1) > 0.005)
            {
                double oldFill = Math.Max(Width / oldAspect, Height);
                double zoom = fill ? 1 : Math.Max(1, PictureHeight / oldFill);
                ph = Math.Max(Width / newAspect, Height) * zoom;
                pw = ph * newAspect;
                ox = fill ? 0 : OffsetX / Math.Max(0.001, PictureWidth) * pw;
                oy = fill ? 0 : OffsetY / Math.Max(0.001, PictureHeight) * ph;
                ox = Math.Max(-(pw-Width)/2, Math.Min((pw-Width)/2, ox));
                oy = Math.Max(-(ph-Height)/2, Math.Min((ph-Height)/2, oy));
            }
            c.PictureWidth = (float)pw;
            c.PictureHeight = (float)ph;
            c.PictureOffsetX = (float)ox;
            c.PictureOffsetY = (float)oy;
            if ((int)s.HorizontalFlip != FlipH) s.Flip(0);
            if ((int)s.VerticalFlip != FlipV) s.Flip(1);
            s.Rotation = Rotation;
            s.Left = Left;
            s.Top = Top;
            s.LockAspectRatio = LockAspect;
        }
    }

    public static class Engine
    {
        public const string IdTag = "IMAGELINK_ASSET_ID";
        public const string NameTag = "IMAGELINK_ASSET_NAME";
        public const string VersionTag = "IMAGELINK_VERSION";
        public static string Tag(dynamic shape, string key) { return (string)shape.Tags.Item(key); }
        public static void Assign(dynamic shape, string id, string name)
        {
            shape.Tags.Add(IdTag, id);
            shape.Tags.Add(NameTag, name);
        }
        public static void Detach(dynamic shape)
        {
            shape.Tags.Delete(IdTag); shape.Tags.Delete(NameTag); shape.Tags.Delete(VersionTag);
        }
        public static List<LinkedPicture> Find(dynamic presentation, string id)
        {
            var result = new List<LinkedPicture>();
            for (int i=1; i <= (int)presentation.Slides.Count; i++)
            {
                dynamic slide = presentation.Slides.Item(i);
                Walk(slide, slide.Shapes, false, id, result);
            }
            return result;
        }
        static void Walk(dynamic slide, dynamic shapes, bool grouped, string id, List<LinkedPicture> result)
        {
            for (int j=1; j <= (int)shapes.Count; j++)
            {
                dynamic s = shapes.Item(j);
                string value = Tag(s, IdTag);
                if (!String.IsNullOrEmpty(value) && (id == null || String.Equals(value,id,StringComparison.OrdinalIgnoreCase)))
                    result.Add(new LinkedPicture { Slide=slide, Shape=s, InGroup=grouped });
                if ((int)s.Type == 6) Walk(slide, s.GroupItems, true, id, result);
            }
        }

        // Validate the actual saved shape XML, including unsupported effects and cross-object references.
        // Restrict v0.1 to ordinary embedded rectangular pictures; never silently flatten complex objects.
        public static void Preflight(string snapshot, List<LinkedPicture> targets)
        {
            if (targets.Count == 0) throw new InvalidOperationException("没有找到关联图片。");
            XNamespace p="http://schemas.openxmlformats.org/presentationml/2006/main";
            XNamespace a="http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace rel="http://schemas.openxmlformats.org/package/2006/relationships";
            using (var zip=ZipFile.OpenRead(snapshot))
            {
                var presentation=ReadXml(zip,"ppt/presentation.xml");
                var rels=ReadXml(zip,"ppt/_rels/presentation.xml.rels");
                var slideIds=presentation.Descendants(p+"sldId").ToList();
                foreach (var target in targets)
                {
                    if (target.InGroup || (int)target.Shape.Type != 13)
                        throw new InvalidOperationException(target.Label + "：第一版仅支持未组合的普通嵌入图片。请先取消组合或解除此处关联。");
                    string rid=(string)slideIds[target.SlideIndex-1].Attribute(r+"id");
                    var relationship=rels.Root.Elements(rel+"Relationship").Single(x=>(string)x.Attribute("Id")==rid);
                    string path=Resolve("ppt/presentation.xml",(string)relationship.Attribute("Target"));
                    var doc=ReadXml(zip,path);
                    string sid=((int)target.Shape.Id).ToString(CultureInfo.InvariantCulture);
                    var nv=doc.Descendants(p+"cNvPr").Single(x=>(string)x.Attribute("id")==sid);
                    var pic=nv.Parent.Parent;
                    bool bad=pic.Name!=p+"pic" || pic.Ancestors(p+"grpSp").Any()
                        || pic.Descendants(p+"ph").Any() || pic.Descendants(a+"custGeom").Any()
                        || pic.Descendants().Any(x=>x.Name.LocalName=="scene3d" || x.Name.LocalName=="sp3d" || x.Name.LocalName=="effectDag")
                        || pic.Descendants().Where(x=>x.Name.LocalName=="extLst").Any(x=>x.Descendants().Any(y=>!y.HasElements && y.Name.LocalName!="creationId" && y.Name.LocalName!="useLocalDpi"))
                        || pic.Descendants(a+"blip").Any(x=>x.Elements().Any(y=>y.Name.LocalName!="extLst") || x.Attribute(r+"link")!=null)
                        || pic.Descendants(a+"effectLst").Elements().Any(x=>x.Name!=a+"outerShdw")
                        || nv.Elements().Any(x=>x.Name.LocalName=="hlinkClick" || x.Name.LocalName=="hlinkHover")
                        || pic.Descendants(a+"picLocks").Attributes().Any(x=>x.Name.LocalName!="noChangeAspect" && x.Value=="1")
                        || doc.Descendants(a+"stCxn").Concat(doc.Descendants(a+"endCxn")).Any(x=>(string)x.Attribute("id")==sid);
                    if (bad) throw new InvalidOperationException(target.Label + "：包含占位符、自定义顶点、特殊效果或交互等暂不支持的设置。本次不会替换任何图片。内置的圆形、圆角矩形等“裁剪为形状”可以正常更新。");
                }
            }
        }
        static XDocument ReadXml(ZipArchive z,string path)
        {
            var entry=z.GetEntry(path); if(entry==null) throw new InvalidDataException("文件结构不完整："+path);
            using(var stream=entry.Open()) return XDocument.Load(stream);
        }
        static string Resolve(string from,string target)
        { return new Uri(new Uri("http://package/"+from),target).AbsolutePath.TrimStart('/'); }

        // A disposable validation snapshot, not a retained backup. Deleted before replacement.
        public static void ValidateCurrentPresentation(dynamic pres, List<LinkedPicture> targets)
        {
            string file=Path.Combine(Path.GetTempPath(),"ImageLink-check-"+Guid.NewGuid().ToString("N")+".pptx");
            try
            {
                pres.SaveCopyAs(file,24,0);
                Preflight(file,targets);
            }
            finally { if(File.Exists(file)) File.Delete(file); }
        }

        public static int Replace(dynamic app, List<LinkedPicture> targets, string file, bool fill)
        {
            var staged=new List<object>();
            bool committing=false;
            try
            {
                app.StartNewUndoEntry();
                foreach(var target in targets)
                {
                    dynamic old=target.Shape;
                    var g=Geometry.Read(old);
                    dynamic fresh=target.Slide.Shapes.AddPicture(file,0,-1,g.Left,g.Top,-1f,-1f);
                    staged.Add((object)fresh);
                    fresh.Visible=0;
                    double aspect=(double)(float)fresh.Width/(float)fresh.Height;
                    old.PickUp(); fresh.Apply();
                    // PowerPoint animations reference a shape object. Copy them to the staged
                    // replacement before the original is removed so ordinary entrance,
                    // emphasis, exit and motion-path effects survive the update.
                    if (HasDirectAnimation(target))
                    {
                        old.PickupAnimation();
                        fresh.ApplyAnimation();
                    }
                    fresh.Visible=0;
                    g.Apply(fresh,aspect,fill);
                    fresh.AlternativeText=(string)old.AlternativeText;
                    fresh.Title=(string)old.Title;
                    for(int t=1;t<=(int)old.Tags.Count;t++) fresh.Tags.Add((string)old.Tags.Name(t),(string)old.Tags.Value(t));
                    fresh.Tags.Add(VersionTag,DateTime.UtcNow.ToString("o"));
                    // Place immediately above the old picture before removing the old one.
                    while((int)fresh.ZOrderPosition > (int)old.ZOrderPosition+1) fresh.ZOrder(3);
                    if(Math.Abs((float)fresh.Width-g.Width)>0.1 || Math.Abs((float)fresh.Height-g.Height)>0.1
                        || Math.Abs((float)fresh.Left-g.Left)>0.1 || Math.Abs((float)fresh.Top-g.Top)>0.1)
                        throw new InvalidOperationException(target.Label+"：未能保留原图片框，已中止更新。");
                }
                // All replacements are prepared successfully before any original is removed.
                committing=true;
                for(int i=0;i<targets.Count;i++)
                {
                    dynamic old=targets[i].Shape; dynamic fresh=staged[i];
                    string name=(string)old.Name; int visible=(int)old.Visible;
                    fresh.Visible=visible;
                    old.Delete(); fresh.Name=name;
                }
                return targets.Count;
            }
            catch(Exception ex)
            {
                if(!committing)
                {
                    foreach(dynamic s in staged) { try { s.Delete(); } catch {} }
                    throw new InvalidOperationException("更新未完成，原图片已保留。\r\n"+ex.Message,ex);
                }
                throw new InvalidOperationException("提交更新时发生错误，可能有部分图片已更新。请先尝试 PowerPoint 的撤销操作；检查结果前不要覆盖已保存的原文件。\r\n"+ex.Message,ex);
            }
        }

        // PickupAnimation raises "Invalid Request" when the shape has no animation.
        // Inspect the slide timeline first and only call it for an actual animation target.
        static bool HasDirectAnimation(LinkedPicture target)
        {
            int shapeId=(int)target.Shape.Id;
            try
            {
                dynamic timeline=target.Slide.TimeLine;
                if (SequenceTargetsShape(timeline.MainSequence,shapeId)) return true;
                dynamic interactive=timeline.InteractiveSequences;
                for(int i=1;i<=(int)interactive.Count;i++)
                    if(SequenceTargetsShape(interactive.Item(i),shapeId)) return true;
            }
            catch { }
            try { return (int)target.Shape.AnimationSettings.Animate != 0; }
            catch { return false; }
        }

        static bool SequenceTargetsShape(dynamic sequence,int shapeId)
        {
            for(int i=1;i<=(int)sequence.Count;i++)
            {
                try { if((int)sequence.Item(i).Shape.Id==shapeId)return true; }
                catch { }
            }
            return false;
        }
    }
}
