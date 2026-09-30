using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ImageLink
{
    [ComVisible(true), Guid("B79D4FE3-C8B2-4B86-91DD-8CE0BBD00414"), ProgId("ImageLink.PowerPointUserAddIn"),
     ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IImageLinkCallbacks))]
    public sealed class AddIn : IDTExtensibility2, IRibbonExtensibility, IImageLinkCallbacks
    {
        dynamic app;
        dynamic ribbon;
        public void OnRibbonLoad(object ui) { ribbon=ui; }
        bool busy;
        readonly NotificationPreferences notifications = new NotificationPreferences(Path.Combine(DataFolder,"success-notifications.txt"));
        public bool GetSuccessNotifications(object control) { return notifications.Enabled; }
        public void SetSuccessNotifications(object control,bool pressed) { Run(()=> { try { notifications.Set(pressed); } finally { if(ribbon!=null) ribbon.Invalidate(); } }); }
        void NotifySuccess(string message) { if(notifications.Enabled) Info(message); }
        public static string DataFolder { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"ImageLinkPowerPointUser"); } }
        public void OnConnection(object application,int mode,object instance,ref Array custom) { app=application; }
        public void OnDisconnection(int mode,ref Array custom) { app=null; }
        public void OnAddInsUpdate(ref Array custom) {}
        public void OnStartupComplete(ref Array custom) {}
        public void OnBeginShutdown(ref Array custom) {}
        public string GetCustomUI(string id)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'><ribbon><tabs>
<tab id='imageLinkTab' label='图片关联'>
 <group id='imageLinkGroups' label='关联设置'>
  <button id='ilCreate' label='建立关联' size='large' imageMso='GroupObjects' onAction='CreateGroup' screentip='为选中的图片建立一个关联组' supertip='建立后可用 Ctrl+C / Ctrl+V 复制到其他页，关联标记会随普通图片复制。'/>
  <button id='ilJoin' label='加入已有关联' size='large' imageMso='HyperlinkInsert' onAction='JoinGroup' screentip='把选中图片加入已有的关联组'/>
  <button id='ilDetach' label='解除关联' size='large' imageMso='HyperlinkRemove' onAction='Detach'/>
  <button id='ilShow' label='查看关联位置' size='large' imageMso='FindDialog' onAction='ShowGroup'/>
 </group>
 <group id='imageLinkUpdate' label='同步更新'>
  <button id='ilUpdate' label='更新关联图片' size='large' imageMso='PictureInsertFromFile' onAction='UpdateGroup' screentip='选择一张新图片，更新当前文档内所有关联位置' supertip='各处独立保留位置、图片框、裁剪、旋转与翻转。更新后请保存文稿。'/>
 </group>
 <group id='imageLinkInfo' label='ImageLink 0.2.0 免管理员版'>
  <button id='ilHelp' label='使用说明' size='large' imageMso='Help' onAction='ShowHelp'/>
  <checkBox id='ilSuccessNotifications' label='显示操作成功通知' getPressed='GetSuccessNotifications' onAction='SetSuccessNotifications' screentip='勾选后显示建立、加入、解除关联及换图成功提示' supertip='默认关闭。设置会在下次打开 PowerPoint 时保留；错误提示和换图确认仍会显示。'/>
 </group>
</tab></tabs></ribbon><contextMenus><contextMenu idMso='ContextMenuPicture'><menu id='picImageLink' label='图片关联' imageMso='GroupObjects'>
<button id='picUpdate' label='更新关联图片' imageMso='PictureInsertFromFile' onAction='UpdateGroup'/>
<button id='picCreate' label='建立关联' imageMso='GroupObjects' onAction='CreateGroup'/>
<button id='picJoin' label='加入已有关联' imageMso='HyperlinkInsert' onAction='JoinGroup'/>
<button id='picDetach' label='解除关联' imageMso='HyperlinkRemove' onAction='Detach'/>
<button id='picShow' label='查看关联位置' imageMso='FindDialog' onAction='ShowGroup'/>
</menu></contextMenu>
<contextMenu idMso='ContextMenuShape'><menu id='shapeImageLink' label='图片关联' imageMso='GroupObjects'>
<button id='shapeUpdate' label='更新关联图片' imageMso='PictureInsertFromFile' onAction='UpdateGroup'/>
<button id='shapeCreate' label='建立关联' imageMso='GroupObjects' onAction='CreateGroup'/>
<button id='shapeJoin' label='加入已有关联' imageMso='HyperlinkInsert' onAction='JoinGroup'/>
<button id='shapeDetach' label='解除关联' imageMso='HyperlinkRemove' onAction='Detach'/>
<button id='shapeShow' label='查看关联位置' imageMso='FindDialog' onAction='ShowGroup'/>
</menu></contextMenu></contextMenus></customUI>";
        }
        void Run(Action action)
        {
            if(busy) return;
            busy=true;
            try { action(); }
            catch(Exception ex)
            {
                MessageBox.Show(Owner(),ex.Message,"图片关联",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }
            finally { busy=false; }
        }
        IWin32Window Owner()
        {
            try { return new WindowHandle(new IntPtr((int)app.HWND)); }
            catch { return null; }
        }
        dynamic Presentation()
        {
            if(app==null || (int)app.Presentations.Count==0) throw new InvalidOperationException("请先打开一份 .pptx 演示文稿。");
            dynamic p=app.ActivePresentation;
            if(String.IsNullOrEmpty((string)p.Path) || !String.Equals(Path.GetExtension((string)p.FullName),".pptx",StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("第一版支持 .pptx。请先将演示文稿另存为 .pptx，再使用此功能。");
            if((int)p.ReadOnly==-1) throw new InvalidOperationException("当前文稿是只读的。请另存为可编辑的 .pptx。");
            return p;
        }
        List<object> Selection(bool one)
        {
            dynamic selection=app.ActiveWindow.Selection;
            if((int)selection.Type!=2) throw new InvalidOperationException("请先选中图片本身（不要处于文字编辑状态）。");
            if((bool)selection.HasChildShapeRange) throw new InvalidOperationException("第一版不支持组合内图片，请先取消组合。");
            dynamic range=selection.ShapeRange;
            if(one && (int)range.Count!=1) throw new InvalidOperationException("请只选中一张关联图片。");
            var result=new List<object>();
            for(int i=1;i<=(int)range.Count;i++)
            {
                dynamic s=range.Item(i);
                if((int)s.Type!=13) throw new InvalidOperationException("请选择普通嵌入图片。第一版不支持组合、图片占位符、外部链接图片或图形填充。");
                result.Add((object)s);
            }
            return result;
        }
        string RequiredId(dynamic shape)
        {
            string id=Engine.Tag(shape,Engine.IdTag);
            if(String.IsNullOrEmpty(id)) throw new InvalidOperationException("这张图片尚未关联。请先点击“建立关联”或“加入已有关联”。");
            return id;
        }
        void Info(string message) { MessageBox.Show(Owner(),message,"图片关联",MessageBoxButtons.OK,MessageBoxIcon.Information); }
        public void CreateGroup(object control) { Run(()=>
        {
            Presentation(); var selected=Selection(false);
            foreach(dynamic s in selected) if(!String.IsNullOrEmpty(Engine.Tag(s,Engine.IdTag)))
                throw new InvalidOperationException("选中图片已有绑定。若要建立新关联，请先解除关联。");
            string name=Dialogs.AskText(Owner(),"建立关联","关联名称，例如：方案 A · 前侧视角");
            if(String.IsNullOrWhiteSpace(name)) return;
            string id=Guid.NewGuid().ToString("N").ToUpperInvariant();
            app.StartNewUndoEntry();
            foreach(dynamic s in selected) Engine.Assign(s,id,name.Trim());
            NotifySuccess("已建立关联："+name+"\r\n\r\n可直接复制粘贴这张图片到其他页，保留关联；已有图片请用“加入已有关联”。\r\n保存 PPT 后关联信息会随文件保存。");
        }); }
        public void JoinGroup(object control) { Run(()=>
        {
            dynamic p=Presentation(); var selected=Selection(false);
            foreach(dynamic s in selected) if(!String.IsNullOrEmpty(Engine.Tag(s,Engine.IdTag)))
                throw new InvalidOperationException("选中图片已有绑定。请先解除关联，再加入其他组。");
            List<LinkedPicture> all=Engine.Find(p,null);
            var groups=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            foreach(var t in all)
            {
                string id=Engine.Tag(t.Shape,Engine.IdTag);
                groups[id]=Engine.Tag(t.Shape,Engine.NameTag)+"   ["+id.Substring(0,Math.Min(6,id.Length))+"]";
            }
            if(groups.Count==0) throw new InvalidOperationException("当前文稿还没有关联组。请先为一张图片建立关联。");
            string picked=Dialogs.PickGroup(Owner(),groups);
            if(picked==null) return;
            var first=all.First(t=>String.Equals(Engine.Tag(t.Shape,Engine.IdTag),picked,StringComparison.OrdinalIgnoreCase));
            app.StartNewUndoEntry();
            foreach(dynamic s in selected) Engine.Assign(s,picked,Engine.Tag(first.Shape,Engine.NameTag));
            NotifySuccess("已加入关联。当前仅建立关系，不改变图片内容；下次点击“更新关联图片”时会一起更新。");
        }); }
        public void Detach(object control) { Run(()=>
        {
            Presentation(); var selected=Selection(false); app.StartNewUndoEntry();
            foreach(dynamic s in selected) Engine.Detach(s);
            NotifySuccess("已解除选中图片的关联。图片内容与排版保持不变。");
        }); }
        public void ShowGroup(object control) { Run(()=>
        {
            dynamic p=Presentation(); dynamic selected=Selection(true)[0];
            List<LinkedPicture> targets=Engine.Find(p,RequiredId(selected));
            Dialogs.ShowLocations(Owner(),Engine.Tag(selected,Engine.NameTag),targets);
        }); }
        public void UpdateGroup(object control) { Run(()=>
        {
            dynamic p=Presentation(); dynamic selected=Selection(true)[0];
            List<LinkedPicture> targets=Engine.Find(p,RequiredId(selected));
            using(var picker=new OpenFileDialog { Title="选择新版本图片", Filter="图片 (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp",CheckFileExists=true,Multiselect=false })
            {
                if(picker.ShowDialog(Owner())!=DialogResult.OK) return;
                using(var img=Image.FromFile(picker.FileName)) { if(img.Width<1 || img.Height<1) throw new InvalidDataException("图片无效。"); }
                bool fill;
                if(!Dialogs.ConfirmUpdate(Owner(),Engine.Tag(selected,Engine.NameTag),targets,picker.FileName,out fill)) return;
                Engine.ValidateCurrentPresentation(p,targets);
                int count=Engine.Replace(app,targets,picker.FileName,fill);
                NotifySuccess("已更新 "+count+" 处关联图片。请检查各页裁剪，再保存 PPT。");
            }
        }); }
        public void ShowHelp(object control) { Run(()=> Info(
            "ImageLink 图片关联 · 0.2.0 免管理员版\r\n\r\n"+
            "1. 保存为 .pptx，选中普通图片，点击“建立关联”。\r\n"+
            "2. Ctrl+C / Ctrl+V 复制到其他页；已有图片使用“加入已有关联”。\r\n"+
            "3. 各处可独立缩放、裁剪、旋转、翻转。\r\n"+
            "4. 选中任意关联图片，用“更新关联图片”选择新图。\r\n"+
            "5. 历史对比图使用“解除关联”，保留当前版本。\r\n\r\n"+
            "请通过本插件换图。PowerPoint 自带“更改图片”不会触发同步。\r\n"+
            "支持普通图片及圆形、圆角矩形等内置“裁剪为形状”；普通动画会随图片更新。组合、占位符、自定义顶点和特殊效果仍会阻止更新。\r\n"+
            "图片比例相同默认保留裁剪；比例变化时等比填满，尽量保留焦点位置。无法自动识别车头等部位。\r\n"+
            "关联信息随正式 PPT 保存，不依赖外部文件。\r\n"+
            "本版不提供自动备份；更新完成后请保存正式文件。\r\n"+
            "操作成功通知默认关闭，可在本栏勾选开启。\r\n\r\n"+
            "不上传图片，不需要网络，不需要 VBA 宏。接收人无需安装插件即可查看 PPT。")); }
    }
    // A per-user preference, independent of presentation data and backups.
    internal sealed class NotificationPreferences
    {
        readonly string path;
        public bool Enabled { get; private set; }
        public NotificationPreferences(string file)
        {
            path=file;
            try { Enabled=File.Exists(path) && File.ReadAllText(path).Trim()=="true"; }
            catch { Enabled=false; }
        }
        public void Set(bool enabled)
        {
            Enabled=enabled;
            try { Directory.CreateDirectory(Path.GetDirectoryName(path)); File.WriteAllText(path,enabled?"true":"false"); }
            catch(Exception ex) { throw new IOException("本次通知设置已生效，但未能保存。下次启动可能恢复旧设置。",ex); }
        }
    }
    sealed class WindowHandle : IWin32Window { public IntPtr Handle {get;private set;} public WindowHandle(IntPtr p){Handle=p;} }
    static class Dialogs
    {
        static Form Base(string title,int height)
        {
            return new Form { Text=title, ClientSize=new Size(620,height), Font=new Font("Microsoft YaHei UI",10),
                StartPosition=FormStartPosition.CenterParent, FormBorderStyle=FormBorderStyle.FixedDialog,
                MaximizeBox=false, MinimizeBox=false, AutoScaleMode=AutoScaleMode.Dpi, ShowInTaskbar=false };
        }
        static void Buttons(Form f,string okay)
        {
            var yes=new Button {Text=okay,DialogResult=DialogResult.OK,Bounds=new Rectangle(380,f.ClientSize.Height-48,105,32)};
            var no=new Button {Text="取消",DialogResult=DialogResult.Cancel,Bounds=new Rectangle(495,f.ClientSize.Height-48,105,32)};
            f.Controls.Add(yes); f.Controls.Add(no); f.AcceptButton=yes; f.CancelButton=no;
        }
        public static string AskText(IWin32Window owner,string title,string caption)
        {
            using(var f=Base(title,165))
            {
                f.Controls.Add(new Label {Text=caption,Bounds=new Rectangle(20,18,580,30)});
                var t=new TextBox {Bounds=new Rectangle(20,55,580,30),MaxLength=80}; f.Controls.Add(t); Buttons(f,"建立");
                return f.ShowDialog(owner)==DialogResult.OK?t.Text:null;
            }
        }
        public static string PickGroup(IWin32Window owner,Dictionary<string,string> groups)
        {
            using(var f=Base("加入已有关联",350))
            {
                f.Controls.Add(new Label {Text="选择要加入的关联组：",Bounds=new Rectangle(20,15,580,30)});
                var list=new ListBox {Bounds=new Rectangle(20,50,580,235),DisplayMember="Value",ValueMember="Key",DataSource=groups.ToList()};
                f.Controls.Add(list); Buttons(f,"加入");
                return f.ShowDialog(owner)==DialogResult.OK?(string)list.SelectedValue:null;
            }
        }
        static ListBox Locations(List<LinkedPicture> targets,int height)
        { return new ListBox { Bounds=new Rectangle(20,60,580,height),DataSource=targets.Select(t=>t.Label+(t.InGroup?"（组合内）":"")).ToList() }; }
        public static void ShowLocations(IWin32Window owner,string name,List<LinkedPicture> targets)
        {
            using(var f=Base("关联位置",400))
            {
                f.Controls.Add(new Label {Text=name+" · 共 "+targets.Count+" 处",Bounds=new Rectangle(20,18,580,35)});
                f.Controls.Add(Locations(targets,275)); Buttons(f,"关闭"); f.ShowDialog(owner);
            }
        }
        public static bool ConfirmUpdate(IWin32Window owner,string name,List<LinkedPicture> targets,string path,out bool fill)
        {
            using(var f=Base("更新关联图片",500))
            {
                f.Controls.Add(new Label {Text=name+" · 将更新 "+targets.Count+" 处",Bounds=new Rectangle(20,15,580,35)});
                f.Controls.Add(Locations(targets,155));
                f.Controls.Add(new Label {Text="新图片："+Path.GetFileName(path),AutoEllipsis=true,Bounds=new Rectangle(20,225,580,30)});
                var keep=new RadioButton {Text="保留各处裁剪（比例变化时适配原图片框）",Checked=true,Bounds=new Rectangle(20,270,580,30)};
                var center=new RadioButton {Text="全部重置为居中填满（保留各处图片框）",Bounds=new Rectangle(20,308,580,30)};
                f.Controls.Add(keep); f.Controls.Add(center);
                f.Controls.Add(new Label {Text="更新后请逐页检查裁剪并保存文稿。\r\n新图中汽车位置变化时，可能需要手动调整局部特写。",Bounds=new Rectangle(20,360,580,60)});
                Buttons(f,"更新全部"); bool ok=f.ShowDialog(owner)==DialogResult.OK; fill=center.Checked; return ok;
            }
        }
    }
}







