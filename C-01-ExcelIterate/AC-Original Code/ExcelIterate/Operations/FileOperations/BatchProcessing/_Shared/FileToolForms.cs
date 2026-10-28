using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Excel_Iterate
{

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


}
