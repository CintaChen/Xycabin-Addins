using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class RenameWorkbookSession
    {
        internal static readonly string[] Headers = { "文档路径", "文件现有名称", "新文件名（待填）", "处理状态" };
        private readonly object app; private readonly string bookPath;
        private string anchor, folder; private List<FileSnapshot> files;
        internal RenameWorkbookSession(object application)
        {
            app = application;
            using (var bag = new ComBag())
            {
                object book = bag.Own(((dynamic)app).ActiveWorkbook);
                if (book == null) throw new InvalidOperationException("请先打开用于接收重命名清单的工作簿。");
                Connect.PrepareWorkbookSave(Convert.ToString(((dynamic)book).FullName));
                bookPath = WorkbookSaving.Save(book, false);
            }
        }
        private object Book(ComBag bag)
        {
            object book = bag.Own(LinkAssistantSession.FindBook(app, bookPath));
            if (book == null) throw new InvalidOperationException("打开窗口时绑定的清单工作簿已关闭，请重新打开窗口。");
            return book;
        }
        private object Sheet(object book, ComBag bag)
        {
            if (anchor == null) throw new InvalidOperationException("请先生成文件清单。");
            dynamic names = bag.Own(((dynamic)book).Names), name = bag.Own(names.Item(anchor)), range = bag.Own(name.RefersToRange);
            return bag.Own(range.Worksheet);
        }
        internal string Create(string root, List<FileSnapshot> scanned)
        {
            using (var bag = new ComBag())
            {
                dynamic book = Book(bag); WorkbookSaving.Save((object)book, false);
                using (var quiet = new HostQuiet(app))
                {
                    object sheet = WorkbookUi.EnsureSheet((object)book, "文件批量命名", bag, false);
                    string token = "_EI_Rename_" + Guid.NewGuid().ToString("N");
                    dynamic names = bag.Own(book.Names); bag.Own(names.Add(token, "=" + WorkbookUi.Quote(Convert.ToString(((dynamic)sheet).Name)) + "!$A$1", false));
                    object[,] values = new object[scanned.Count + 1, 4];
                    for (int c = 0; c < 4; c++) values[0,c] = Headers[c];
                    for (int i = 0; i < scanned.Count; i++) { values[i+1,0] = scanned[i].Path; values[i+1,1] = scanned[i].Name; values[i+1,2] = ""; values[i+1,3] = "待填写"; }
                    dynamic range = bag.Own(((dynamic)sheet).Range["A1:D" + (scanned.Count+1)]); range.NumberFormat = "@"; range.Value2 = values;
                    WorkbookUi.StyleTable(sheet, 4, scanned.Count + 1);
                    WorkbookUi.Center(range);
                    WorkbookUi.FitOutput(sheet, 1, 4, scanned.Count + 1);
                    ((dynamic)sheet).Activate(); folder = root; files = scanned; anchor = token;
                    WorkbookSaving.Save((object)book, false);
                    return "已在绑定工作簿中生成「" + Convert.ToString(((dynamic)sheet).Name) + "」，共 " + files.Count + " 个文件。填写第三列后点击执行替换。";
                }
            }
        }
        internal RenamePlan Read()
        {
            using (var bag = new ComBag())
            {
                object book = Book(bag); Connect.PrepareWorkbookSave(bookPath); WorkbookSaving.Save(book, false);
                dynamic sheet = Sheet(book, bag), used = bag.Own(sheet.UsedRange), rows = bag.Own(used.Rows);
                if (Convert.ToInt32(rows.Count) != files.Count + 1) throw new InvalidOperationException("清单行数已变化，请重新生成清单；不要增删行。");
                dynamic range = bag.Own(sheet.Range["A1:D" + (files.Count+1)]); Array data = range.Value2 as Array;
                var plan=Parse(folder, files, data);plan.DocumentPath=bookPath;plan.Anchor=anchor;return plan;
            }
        }
        internal static RenamePlan Parse(string root, IList<FileSnapshot> knownFiles, Array data)
        {
            if (data == null || data.Rank != 2 || data.GetLength(0) != knownFiles.Count+1 || data.GetLength(1) != 4) throw new InvalidOperationException("无法读取完整清单，或行数已变化。");
            int r0 = data.GetLowerBound(0), c0 = data.GetLowerBound(1);
            for (int c = 0; c < 4; c++) if (Convert.ToString(data.GetValue(r0,c0+c)) != Headers[c]) throw new InvalidOperationException("清单表头已变化，请重新生成清单。");
            var known = knownFiles.ToDictionary(f => f.Path, StringComparer.OrdinalIgnoreCase); var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase); var plan = new RenamePlan { Root = root };
            for (int r = 1; r <= knownFiles.Count; r++)
            {
                string path = Convert.ToString(data.GetValue(r0+r,c0)); FileSnapshot file;
                if (!known.TryGetValue(path, out file) || !seen.Add(path) || path != file.Path || Convert.ToString(data.GetValue(r0+r,c0+1)) != file.Name) throw new InvalidOperationException("原路径或原文件名已变化，或有重复行；请重新生成清单。");
                plan.Items.Add(new RenameItem { Source = file, Input = Convert.ToString(data.GetValue(r0+r,c0+2)) });
            }
            return plan;
        }
        internal void Restore(OperationBatch batch)
        {
            if(batch.Kind!="Rename" || !string.Equals(batch.DocumentPath,bookPath,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("请打开原清单工作簿后恢复该批次。");
            folder=batch.Root;anchor=batch.Anchor;files=batch.ReportFiles.Select(f=>f.Snapshot()).ToList();
        }
        internal void WriteResult(FileOperationResult result)
        {
            using (var bag = new ComBag())
            {
                object book = Book(bag); dynamic sheet = Sheet(book, bag), used = bag.Own(sheet.UsedRange), rows = bag.Own(used.Rows);
                if (Convert.ToInt32(rows.Count) != files.Count+1) throw new InvalidOperationException("处理期间清单行数已变化，文件结果请以日志为准。");
                dynamic range = bag.Own(sheet.Range["A1:D" + (files.Count+1)]);
                Array data = range.Value2 as Array; var plan = Parse(folder, files, data);
                for (int i = 0; i < plan.Items.Count; i++)
                {
                    string status; if (!result.Status.TryGetValue(plan.Items[i].Source.Path, out status)) status = result.Failed > 0 ? "整批暂停，未执行" : "未执行";
                    dynamic cell = bag.Own(sheet.Range["D" + (i+2)]); cell.Value2 = status;
                }
                WorkbookUi.FitOutput((object)sheet, 1, 4, files.Count + 1);
                sheet.Activate(); WorkbookSaving.Save(book, false);
            }
        }
    }
    internal class FileToolForm : Form, IDialogDefaultLayout
    {
        internal readonly TextBox FolderPath=new TextBox { Anchor=AnchorStyles.Left|AnchorStyles.Right };
        internal readonly RichTextBox LogView=new RichTextBox { ReadOnly=true,WordWrap=true,ScrollBars=RichTextBoxScrollBars.Vertical,Dock=DockStyle.Fill,BorderStyle=BorderStyle.None,BackColor=Color.White };
        internal readonly Label Status=Ui.Label("请选择文件夹，或粘贴带引号的完整路径。",9,false);
        internal readonly FlowLayoutPanel Options=new FlowLayoutPanel { Dock=DockStyle.Top,AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,WrapContents=false,FlowDirection=FlowDirection.TopDown,Margin=new Padding(0,0,0,12) };
        internal readonly FlowLayoutPanel Actions=Ui.ActionBar();
        internal readonly Button Browse=Ui.Button("浏览…",90),Stop=Ui.Button("取消操作",104),Confirm=Ui.Button("确认操作",104);
        internal readonly TableLayoutPanel Layout;
        protected bool Busy;protected volatile bool CancelRequested;
        internal bool IsWorking {get{return Busy;}}
        protected OperationBatch Batch;
        protected string Introduction="";
        private bool closeAfter,loading,sizing;
        private readonly TableLayoutPanel path;
        internal FileToolForm(string title,string windowId)
        {
            Text=title;DialogSizing.Prepare(this,windowId,"document",Ui.Blue);KeyPreview=true;
            Layout=new TableLayoutPanel { Dock=DockStyle.Fill,ColumnCount=1,RowCount=5,Padding=new Padding(16) };
            Layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            Layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));Layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));Layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));Layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));Layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            path=new TableLayoutPanel { Dock=DockStyle.Top,AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,ColumnCount=2,RowCount=1,Margin=new Padding(0,0,0,12) };
            path.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            path.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));path.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));path.Controls.Add(FolderPath,0,0);path.Controls.Add(Browse,1,0);
            Browse.Click+=delegate {string selected=HostDialogService.Folder(this,"选择需要处理的文件夹",false);if(selected!=null)FolderPath.Text=selected;};
            Status.AutoSize=true;Status.AutoEllipsis=false;Status.MaximumSize=new Size(760,0);Status.Margin=new Padding(0,8,0,12);
            Stop.Click+=delegate {if(Busy){CancelRequested=true;Status.Text="等待当前步骤完成，然后还原本批改动…";}else if(Batch!=null&&Batch.CanUndo)Run(()=>Batch.Undo(null),UndoFinished);};
            Confirm.Click+=delegate {
                if(Busy||Batch==null)return;
                if(UiDialog.Show(this,Batch.ConfirmationText(),"确认操作",MessageBoxButtons.OKCancel,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2)!=DialogResult.OK)return;
                Run(delegate {Batch.Confirm(RecycleFiles.Send);return true;},delegate(bool done){Batch=null;Status.Text="已确认；本批日志和备份已清理。";SetLog(Status.Text);});
            };
            var copy=Ui.Button("复制日志",104);var close=Ui.Button("关闭",90);
            copy.Click+=delegate {try {Clipboard.SetText(LogView.Text.Length==0?Status.Text:LogView.Text);}catch(Exception ex){Status.Text="复制失败："+ex.Message;}};
            close.Click+=delegate {Close();};Actions.Controls.AddRange(new Control[]{Stop,Confirm,copy,close});
            var card=new UiCard { Dock=DockStyle.Fill,Padding=new Padding(12),MinimumSize=new Size(0,100),Margin=new Padding(0,0,0,12) };card.Controls.Add(LogView);
            Layout.Controls.Add(path,0,0);Layout.Controls.Add(Status,0,1);Layout.Controls.Add(Options,0,2);Layout.Controls.Add(card,0,3);Layout.Controls.Add(Actions,0,4);Controls.Add(Layout);
            FolderPath.TextChanged+=delegate {if(Busy||loading||DialogPreviewFactory.IsPreview(this))return;try {string root=FilePaths.Folder(FolderPath.Text);var found=OperationBatch.Find(root);if(found!=null){loading=true;Batch=found;RestoreBatch(found);Status.Text="已恢复待确认批次；可取消还原或确认清理。";SetLog(File.Exists(found.LogPath)?File.ReadAllText(found.LogPath):found.Journal);} }catch(Exception ex){if(Directory.Exists(FolderPath.Text.Trim('"')))Status.Text=ex.Message;}finally{loading=false;RefreshActions();}};
            FormClosing+=delegate(object sender,FormClosingEventArgs e){if(Busy){CancelRequested=true;closeAfter=true;e.Cancel=true;Status.Text="等待安全步骤结束，保留记录后关闭。";}};
            Status.TextChanged+=delegate {Status.Visible=Status.Text.Length>0;};
            Resize+=delegate {SizeRows();};
            Shown+=delegate {
                SizeRows();Layout.PerformLayout();
            };RefreshActions();
        }
        public Size MeasureDefaultClient()
        {
            SizeRows();Layout.PerformLayout();int used=Layout.Padding.Vertical;
            // Visible reports false while the measurement form is hidden. These
            // initial sections are selected by content, never by parent visibility.
            foreach(Control item in new Control[]{path,Status,Options,Actions})
                if(item!=Status&&item!=Options||item==Status&&Status.Text.Length>0||item==Options&&Options.Controls.Count>0)
                    used+=item.GetPreferredSize(new Size(Math.Max(100,Layout.ClientSize.Width-Layout.Padding.Horizontal),0)).Height+item.Margin.Vertical;
            int height=used+Ui.Px(this,160)+Ui.Px(this,24)+Ui.Px(this,12);
            return new Size(DialogSizing.LogicalClient(this).Width,(int)Math.Ceiling(height/DialogSizing.Scale(this)));
        }
        protected virtual void SizeOptionRows(int width) { }
        private void SizeRows()
        {
            if(sizing||Layout==null)return;sizing=true;
            try {int width=Math.Max(100,ClientSize.Width-Layout.Padding.Horizontal);Options.Visible=Options.Controls.Count>0;Options.MaximumSize=new Size(width,0);Options.Width=width;path.MaximumSize=new Size(width,0);Status.MaximumSize=new Size(width,0);Actions.MaximumSize=new Size(width,0);SizeOptionRows(width);}
            finally{sizing=false;}
        }
        protected virtual void RestoreBatch(OperationBatch batch) { }
        protected virtual void UndoFinished(FileOperationResult result) {Result(result);}
        protected bool Ready()
        {if(Busy||DialogPreviewFactory.IsPreview(this))return false;if(Batch!=null){Status.Text="有待确认批次，请先取消还原或确认清理。";return false;}return true;}
        protected void SetLog(string value)
        {
            LogView.Clear();if(Introduction.Length>0){LogView.SelectionFont=Ui.MakeFont(9,FontStyle.Bold,Font);LogView.AppendText(Introduction+"\r\n\r\n");}
            LogView.SelectionFont=Ui.MakeFont(9,FontStyle.Regular,Font);LogView.AppendText(value);
        }
        protected void RefreshActions()
        {
            FolderPath.Enabled=Browse.Enabled=Options.Enabled=!Busy&&Batch==null;
            foreach(Control b in Actions.Controls)b.Enabled=!Busy&&Batch==null||b.Text=="复制日志"||b.Text=="关闭";
            Stop.Enabled=Busy||Batch!=null&&Batch.CanUndo;
            Confirm.Enabled=!Busy&&Batch!=null&&Batch.State!="UndoFailed";
        }
        protected void Post(Action action){if(!IsDisposed&&IsHandleCreated)BeginInvoke(action);}
        protected void Progress(string text){Post(delegate {Status.Text=text;});}
        protected void Run<T>(Func<T> work,Action<T> finished)
        {
            if(Busy)return;Busy=true;CancelRequested=false;RefreshActions();
            var worker=new Thread(delegate(){T result=default(T);Exception error=null;try{result=work();}catch(Exception ex){error=ex;}
                Post(delegate {Busy=false;try {if(error==null)finished(result);else {Status.Text="未完成："+error.Message;LogView.AppendText("\r\n"+error);}}catch(Exception ex){Status.Text="后续处理失败："+ex.Message;LogView.AppendText("\r\n"+ex);}RefreshActions();if(closeAfter&&!Busy)Close();});});
            worker.SetApartmentState(ApartmentState.STA);worker.IsBackground=true;worker.Start();
        }
        protected void Result(FileOperationResult result)
        {
            if(result.Batch!=null)Batch=result.Batch;Status.Text=result.Summary;SetLog(result.Text+(Batch==null?"":"\r\n日志："+Batch.LogPath));RefreshActions();
        }
        protected override void OnKeyDown(KeyEventArgs e){if(e.KeyCode==Keys.Escape){e.Handled=true;Close();}base.OnKeyDown(e);}
    }
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
    internal sealed class RenameFilesForm : FileToolForm
    {
        private readonly RenameWorkbookSession session;
        internal readonly Button Scan=Ui.Button("生成文件清单",140),Execute=Ui.Button("执行替换",112);
        internal RenameFilesForm(RenameWorkbookSession bound):base("文件批量命名","file.rename")
        {
            session=bound;Introduction="1. 请浏览选择文件夹，或自行粘贴带引号的完整路径。\r\n2. 只读取第一层文件；第三列可只填文件名或完整文件名（文件名+扩展名）。";SetLog("");Status.Text="";
            Actions.Controls.Add(Scan);Actions.Controls.SetChildIndex(Scan,0);Actions.Controls.Add(Execute);Actions.Controls.SetChildIndex(Execute,1);
            Scan.Click+=delegate {if(!Ready())return;string root;try{root=FilePaths.Folder(FolderPath.Text);}catch(Exception ex){Status.Text=ex.Message;return;}
                Run(delegate {var notes=new List<string>();return FilePaths.Scan(root,false,notes);},delegate(List<FileSnapshot> files){if(CancelRequested){Status.Text="已取消，未生成清单。";return;}if(session==null){Status.Text="预览窗口不连接工作簿。";return;}Status.Text=session.Create(root,files);SetLog(Status.Text);});};
            Execute.Click+=delegate {if(!Ready())return;RenamePlan plan;try {if(session==null)throw new InvalidOperationException("预览窗口不连接工作簿。");plan=session.Read();}catch(Exception ex){Status.Text=ex.Message;LogView.AppendText("\r\n"+ex.Message);return;}
                Run(()=>FileRenamer.Execute(plan,()=>CancelRequested,Progress,null),delegate(FileOperationResult result){WriteStatus(result);Result(result);});};
        }
        protected override void RestoreBatch(OperationBatch batch){if(session!=null)session.Restore(batch);}
        protected override void UndoFinished(FileOperationResult result){WriteStatus(result);Result(result);}
        private void WriteStatus(FileOperationResult result)
        {try {if(session!=null)session.WriteResult(result);result.Log.AppendLine("清单状态已回写并自动保存。");}catch(Exception ex){result.Log.AppendLine("清单回写失败，文件实际结果仍有效："+ex.Message);}if(result.Batch!=null)result.Batch.Log(result.Text);}
    }
}
