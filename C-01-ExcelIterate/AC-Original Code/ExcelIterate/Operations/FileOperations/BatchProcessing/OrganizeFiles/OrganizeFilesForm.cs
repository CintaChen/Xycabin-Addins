using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class OrganizeFilesForm : FileToolForm
    {
        internal readonly CheckBox Replace=new CheckBox { Text="替换同名文件（保留最新修改）",AutoSize=true },Classify=new CheckBox { Text="按扩展名分类文件夹",AutoSize=true },Filter=new CheckBox { Text="仅保留指定类型（其余暂存到备份）",AutoSize=true };
        internal readonly FlowLayoutPanel Types=new FlowLayoutPanel { AutoScroll=true,WrapContents=false,FlowDirection=FlowDirection.LeftToRight,Height=34,Width=740,BackColor=Color.White,Margin=Padding.Empty };
        internal readonly Label CustomTitle=Ui.Label("自定义文件批量整理选项",10,true);
        private readonly FlowLayoutPanel row1=new FlowLayoutPanel { AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,WrapContents=true },row2=new FlowLayoutPanel { AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,WrapContents=true };
        internal readonly Button ReadTypes=Ui.Button("读取文件类型",130),Preview=Ui.Button("执行整理",112);
        internal OrganizeFilesForm():base("文件批量整理","file.organize")
        {
            CustomTitle.AutoSize=true;CustomTitle.Dock=DockStyle.None;CustomTitle.ForeColor=Ui.Blue;
            row1.Controls.AddRange(new Control[]{Replace,Classify});row2.Controls.AddRange(new Control[]{Filter,ReadTypes});Options.Controls.AddRange(new Control[]{CustomTitle,row1,row2,Types});
            SetTypes(new[]{"xlsx","xls","docx","doc","pdf","txt",""});Types.Enabled=false;
            Filter.CheckedChanged+=delegate {Types.Enabled=Filter.Checked;};
            Types.SizeChanged+=delegate {SizeOptionRows(Options.Width);};
            ReadTypes.Click+=delegate {if(!Ready())return;string root;try {root=FilePaths.Folder(FolderPath.Text);}catch(Exception ex){Status.Text=ex.Message;return;}
                Run(delegate {var notes=new List<string>();return FilePaths.Scan(root,true,notes).Select(f=>FilePaths.Extension(f.Path)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(t=>t).ToArray();},delegate(string[] found){SetTypes(found);Types.Enabled=Filter.Checked;Status.Text=found.Length==0?"未发现可选文件类型。":"文件类型已读取，可勾选多种类型。";});};
            Actions.Controls.Add(Preview);Actions.Controls.SetChildIndex(Preview,0);
            Preview.Click+=delegate {if(!Ready())return;var options=new OrganizeOptions { Replace=Replace.Checked,Classify=Classify.Checked,Filter=Filter.Checked };
                foreach(CheckBox c in Types.Controls)if(c.Checked)options.Types.Add((string)c.Tag);string root=FolderPath.Text;
                Run(()=>FileOrganizer.Plan(root,options),delegate(OrganizePlan plan){SetLog(plan.Preview);Status.Text="请核对整理清单。";
                    if(CancelRequested||UiDialog.Show(this,plan.Confirmation,"确认文件批量整理",MessageBoxButtons.OKCancel,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2)!=DialogResult.OK){Status.Text="已取消确认，未处理文件。";return;}
                    Run(()=>FileOrganizer.Execute(plan,null,()=>CancelRequested,Progress,null),Result);});};
        }
        internal void SetPreviewTypes(){SetTypes(new[]{"xlsx","docx","pdf"});}
        private void SetTypes(IEnumerable<string> types)
        {foreach(Control c in Types.Controls.Cast<Control>().ToArray())c.Dispose();Types.Controls.Clear();foreach(string type in types)Types.Controls.Add(new CheckBox { Text=type.Length==0?"无扩展名":type,Tag=type,AutoSize=true,Margin=new Padding(6,4,12,4) });Types.Visible=Types.Controls.Count>0;SizeOptionRows(Options.Width);}
        protected override void SizeOptionRows(int width)
        {
            if(Types==null||row1==null)return;int w=Math.Max(100,width-Ui.Px(this,8));row1.MaximumSize=row2.MaximumSize=new Size(w,0);row1.Width=row2.Width=w;Types.Width=w;
            int line=Types.Controls.Count==0?Ui.Px(this,24):Types.Controls.Cast<Control>().Max(c=>c.GetPreferredSize(Size.Empty).Height+c.Margin.Vertical);
            int total=Types.Controls.Cast<Control>().Sum(c=>c.GetPreferredSize(Size.Empty).Width+c.Margin.Horizontal);
            int h=line+(total>Types.ClientSize.Width?SystemInformation.HorizontalScrollBarHeight:0);if(Types.Height!=h)Types.Height=h;
        }
    }
}
