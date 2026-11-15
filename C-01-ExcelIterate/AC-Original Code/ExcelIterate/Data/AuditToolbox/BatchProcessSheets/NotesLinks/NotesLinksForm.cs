using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class NotesLinksForm : Form
    {
        internal object Application;
        internal Func<bool> CanRun;
        internal event Action<bool> BusyChanged;
        private readonly TextBox parent=Box("",480),folder=Box("",480),parentCode=Box("K2",65),codeColumn=Box("K",45),nameColumn=Box("M",45),fileColumn=Box("N",45),sourceRange=Box("D1:X1700",155),firstStart=Box("AV1",75),secondStart=Box("",75),firstColumn=Box("D",45),lastColumn=Box("X",45);
        private readonly ComboBox leadSheet=Choice("Lead"),sourceSheet=Choice("Notes"),targetSheet=Choice("Notes");
        private readonly NumericUpDown firstRow=Number(3,1),lastRow=Number(0,0),sourceFirstRow=Number(1,1,1048576),sourceRows=Number(1700,1);
        private readonly FlowLayoutPanel settings=new FlowLayoutPanel{Dock=DockStyle.Fill,FlowDirection=FlowDirection.TopDown,WrapContents=false,AutoScroll=true};
        private readonly DataGridView grid=new ResultGrid{Dock=DockStyle.Fill,AllowUserToAddRows=false,AllowUserToDeleteRows=false,RowHeadersVisible=false};
        private readonly RichTextBox log=new RichTextBox{Dock=DockStyle.Fill,ReadOnly=true,BackColor=Color.White};
        private readonly ProgressBar progress=new ProgressBar{Dock=DockStyle.Fill,Style=ProgressBarStyle.Continuous};
        private readonly Label state=new Label{AutoSize=true,Text="选择父公司，读取名单，再预检。结束行0表示自动。"};
        private readonly RoundedButton read=Ui.Button("读取公司",115),preview=Ui.Button("预检查",100),execute=Ui.Button("备份后更新",130),cancel=Ui.Button("停止",90),undo=Ui.Button("还原本次更新",135);
        private readonly Timer timer=new Timer{Interval=25};
        private List<NotesLinkTask> tasks=new List<NotesLinkTask>();
        private NotesLinkPlan plan;
        private NotesLinkExecution execution;
        private IEnumerator<string> sequence;
        private string leadKey,leadHash,lastParent,lastBackup,lastHash,sheetsPath;
        private bool loading,busy,ticking,closePending;
        private volatile bool stop;
        private static TextBox Box(string value,int width){return new TextBox{Text=value,Width=width};}
        private static ComboBox Choice(string value){return new ComboBox{Text=value,Width=160,DropDownStyle=ComboBoxStyle.DropDown};}
        private static NumericUpDown Number(int value,int minimum,int maximum=100000){return new NumericUpDown{Minimum=minimum,Maximum=maximum,Value=value,Width=82};}
        private static Label Label(string text){return new Label{Text=text,AutoSize=true,Anchor=AnchorStyles.Left,Margin=new Padding(4,7,5,3)};}
        internal NotesLinksForm()
        {
            Text="Notes 批量链接";DialogSizing.Prepare(this,"notes.links","table",Ui.Blue);
            var root=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=6,Padding=new Padding(14)};root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,400));root.RowStyles.Add(new RowStyle(SizeType.Percent,60));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));root.RowStyles.Add(new RowStyle(SizeType.Absolute,18));root.RowStyles.Add(new RowStyle(SizeType.Percent,40));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var chooseParent=Ui.Button("选择父公司",125);chooseParent.Click+=delegate{Attempt(delegate{var paths=HostDialogService.OpenFiles(this,"Excel工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm");if(paths.Length==1){parent.Text=paths[0];if(folder.Text.Length==0)folder.Text=Path.GetDirectoryName(paths[0]);RefreshSheets();}});};
            var chooseFolder=Ui.Button("选择文件夹",125);chooseFolder.Click+=delegate{string path=HostDialogService.Folder(this,"子公司工作簿所在目录",true);if(!string.IsNullOrEmpty(path))folder.Text=path;};
            Row("父公司",parent,chooseParent);Row("子公司目录",folder,chooseFolder);
            Row("单体 Notes",sourceSheet,SheetPick(sourceSheet,SampleFile));Row("父公司 Notes",targetSheet,SheetPick(targetSheet,delegate{return parent.Text;}));
            Row("来源区域",sourceRange,Pick("来源",SampleFile,SampleSheet,delegate(MigrationArea a){sourceRange.Text=a.Address;SyncRange();}));
            Row("第一粘贴起点",firstStart,Pick("第一起点",delegate{return parent.Text;},delegate{return targetSheet.Text;},delegate(MigrationArea a){TargetAnchor(a);firstStart.Text=CellAddress.Name(a.Row,a.Column);}));
            Row("第二粘贴起点",secondStart,Pick("第二起点",delegate{return parent.Text;},delegate{return targetSheet.Text;},delegate(MigrationArea a){TargetAnchor(a);secondStart.Text=CellAddress.Name(a.Row,a.Column);}));
            var advanced=Ui.Button("高级设置 ▼",150);Row("",advanced);int advancedStart=settings.Controls.Count;
            Row("Lead",leadSheet);
            Row("Lead字段",Label("父编码"),parentCode,Pick("父编码",delegate{return parent.Text;},delegate{return leadSheet.Text;},delegate(MigrationArea a){Single(a);parentCode.Text=CellAddress.Name(a.Row,a.Column);}),Label("编码列"),codeColumn,ColumnPick(codeColumn),Label("全称列"),nameColumn,ColumnPick(nameColumn),Label("完整文件名列"),fileColumn,ColumnPick(fileColumn));
            Row("名单行",Label("起始"),firstRow,Label("结束（0自动）"),lastRow,Pick("名单行",delegate{return parent.Text;},delegate{return leadSheet.Text;},delegate(MigrationArea a){firstRow.Value=a.Row;lastRow.Value=a.Row+a.Rows-1;}),Label("默认N列为带编码的完整文件名，不使用O列的无编码文件名。"));
            Row("复制行数",sourceRows);
            var advancedRows=settings.Controls.Cast<Control>().Skip(advancedStart).ToArray();foreach(var item in advancedRows)item.Visible=false;advanced.Click+=delegate{bool expand=advanced.Text.EndsWith("▼");foreach(var item in advancedRows)item.Visible=expand;advanced.Text=expand?"高级设置 ▲":"高级设置 ▼";};
            Row("规则",Label("取消勾选保留原槽位；非标准表名或未找到横跨来源列的合并区域时，检查后勾选“确认例外”。零显示为-；不取消合并，不改边框及尺寸。"));
            grid.Columns.Add(new DataGridViewCheckBoxColumn{Name="selected",HeaderText="执行",Width=48});grid.Columns.Add(new DataGridViewTextBoxColumn{Name="code",HeaderText="编码",Width=85,ReadOnly=true});grid.Columns.Add(new DataGridViewTextBoxColumn{Name="company",HeaderText="公司",Width=160,ReadOnly=true});
            grid.Columns.Add(new DataGridViewTextBoxColumn{Name="file",HeaderText="来源文件（可修改）",Width=220});grid.Columns.Add(new DataGridViewTextBoxColumn{Name="sheet",HeaderText="来源工作表",Width=140});grid.Columns.Add(new DataGridViewCheckBoxColumn{Name="confirmed",HeaderText="确认例外",Width=75});grid.Columns.Add(new DataGridViewTextBoxColumn{Name="target",HeaderText="父公司目标区域",Width=160,ReadOnly=true});grid.Columns.Add(new DataGridViewTextBoxColumn{Name="status",HeaderText="状态",AutoSizeMode=DataGridViewAutoSizeColumnMode.Fill,ReadOnly=true});Ui.StyleGrid(grid);
            var chooseSource=Ui.Button("指定所选来源文件",170);chooseSource.Click+=delegate{Attempt(ChangeFile);};var all=Ui.Button("全选",85);var clear=Ui.Button("取消全选",100);all.Click+=delegate{SelectAll(true);};clear.Click+=delegate{SelectAll(false);};
            var actions=Ui.ActionBar();var close=Ui.Button("关闭",90);actions.Controls.AddRange(new Control[]{read,chooseSource,all,clear,preview,execute,cancel,undo,close});
            root.Controls.Add(settings,0,0);root.Controls.Add(grid,0,1);root.Controls.Add(state,0,2);root.Controls.Add(progress,0,3);root.Controls.Add(log,0,4);root.Controls.Add(actions,0,5);Controls.Add(root);
            parent.Leave+=delegate{if(!busy&&File.Exists(parent.Text)&&parent.Text!=sheetsPath)Attempt(RefreshSheets);};
            read.Click+=delegate{Attempt(ReadCompanies);};preview.Click+=delegate{Preview();};execute.Click+=delegate{Attempt(Begin);};cancel.Click+=delegate{stop=true;if(execution!=null)execution.Cancelled=true;Append("已请求停止；提交前取消将丢弃临时修改。");};undo.Click+=delegate{Attempt(Restore);};close.Click+=delegate{Close();};execute.Enabled=cancel.Enabled=undo.Enabled=false;
            foreach(Control control in new Control[]{parent,folder,leadSheet,sourceSheet,targetSheet,parentCode,codeColumn,nameColumn,fileColumn,firstRow,lastRow,sourceRange,firstStart,secondStart})
            {if(control is NumericUpDown)((NumericUpDown)control).ValueChanged+=delegate{Changed();};else control.TextChanged+=delegate{Changed();};}
            sourceSheet.TextChanged+=delegate{if(!loading){foreach(var task in tasks){task.Sheet="";task.NameNeedsConfirmation=false;}Render();}};
            sourceRange.TextChanged+=delegate{if(!loading)try{SyncRange();}catch(InvalidOperationException){};};sourceRange.Leave+=delegate{Attempt(SyncRange);};firstColumn.TextChanged+=delegate{Parameters();};lastColumn.TextChanged+=delegate{Parameters();};sourceFirstRow.ValueChanged+=delegate{Parameters();};sourceRows.ValueChanged+=delegate{Parameters();};
            grid.CurrentCellDirtyStateChanged+=delegate{if(grid.IsCurrentCellDirty)grid.CommitEdit(DataGridViewDataErrorContexts.Commit);};grid.CellValueChanged+=Edited;grid.DataError+=delegate(object sender,DataGridViewDataErrorEventArgs e){e.ThrowException=false;};
            timer.Tick+=delegate{Tick();};FormClosing+=delegate(object sender,FormClosingEventArgs e){if(busy){e.Cancel=true;closePending=true;stop=true;if(execution!=null)execution.Cancelled=true;}};
            Disposed+=delegate{timer.Stop();timer.Dispose();if(sequence!=null)sequence.Dispose();if(execution!=null)execution.Dispose();};
        }
        private void Row(string name,params Control[] fields)
        {var line=new FlowLayoutPanel{AutoSize=true,WrapContents=true,Margin=new Padding(0,1,0,1),Width=1120};line.Controls.Add(new Label{Text=name,AutoSize=false,Width=90,Height=28,TextAlign=ContentAlignment.MiddleLeft});line.Controls.AddRange(fields);settings.Controls.Add(line);}
        private void Append(string text){if(IsDisposed)return;log.AppendText(DateTime.Now.ToString("HH:mm:ss")+"  "+text+Environment.NewLine);log.SelectionStart=log.TextLength;log.ScrollToCaret();log.Update();}
        private void Attempt(Action action){try{if(busy||CanRun!=null&&!CanRun())throw new InvalidOperationException("请等待当前任务结束。");action();}catch(OperationCanceledException error){Append(error.Message);}catch(Exception error){Append("失败："+error.Message);UiDialog.Show(this,error.Message,"Notes 批量链接",MessageBoxButtons.OK,MessageBoxIcon.Warning);}}
        private void Changed(){if(loading)return;plan=null;execute.Enabled=false;state.Text="参数或名单已改变，请重新预检。";}
        private void SetBusy(bool value)
        {busy=value;settings.Enabled=grid.Enabled=read.Enabled=preview.Enabled=!value;cancel.Enabled=value;execute.Enabled=!value&&plan!=null&&plan.Tasks.All(t=>t.Status.StartsWith("通过",StringComparison.Ordinal));undo.Enabled=!value&&lastBackup!=null;if(BusyChanged!=null)BusyChanged(value);}
        private NotesLinkConfig Config()
        {return new NotesLinkConfig{Parent=parent.Text.Trim().Trim('"'),Folder=folder.Text.Trim().Trim('"'),LeadSheet=leadSheet.Text,ParentCodeCell=parentCode.Text.Trim(),CodeColumn=codeColumn.Text.Trim(),NameColumn=nameColumn.Text.Trim(),FileColumn=fileColumn.Text.Trim(),FirstRow=(int)firstRow.Value,LastRow=(int)lastRow.Value,SourceSheet=sourceSheet.Text,TargetSheet=targetSheet.Text,SourceRange=sourceRange.Text.Trim(),FirstStart=firstStart.Text.Trim(),SecondStart=secondStart.Text.Trim()};}
        private void RefreshSheets()
        {
            var names=DataMigrationHost.PackageSheets(parent.Text);loading=true;
            try{SetChoices(leadSheet,names,leadSheet.Text);SetChoices(targetSheet,names,targetSheet.Text);sheetsPath=parent.Text;}finally{loading=false;}Changed();
        }
        private static void SetChoices(ComboBox box,List<string> names,string preferred)
        {box.Items.Clear();box.Items.AddRange(names.Cast<object>().ToArray());var matches=NotesLinkRules.Candidates(names,preferred);box.Text=matches.Count==1&&NotesLinkRules.ExactOrAlias(matches[0],preferred)?matches[0]:preferred;}
        private void ReadCompanies()
        {
            var cfg=Config();NotesLinkExecution.CheckOpenWorkbooks(Application,new string[0],cfg.Parent,false);tasks=NotesLinkRules.ReadLead(cfg);leadKey=cfg.LeadKey;leadHash=DataMigrationHost.Digest(cfg.Parent);Changed();
            var sample=tasks.FirstOrDefault(t=>File.Exists(t.File));if(sample!=null){loading=true;try{SetChoices(sourceSheet,DataMigrationHost.PackageSheets(sample.File),sourceSheet.Text);}finally{loading=false;}}
            Render();Append("已读取直系公司 "+tasks.Count+" 家，包含名单中的本部，按Lead顺序固定槽位；父编码来源："+cfg.LeadSheet+"!"+cfg.ParentCodeCell);
        }
        private void Render()
        {
            loading=true;
            try
            {
                grid.Rows.Clear();var cfg=Config();
                foreach(var task in tasks)
                {
                    string target="待填写两个起点";try{target=NotesLinkRules.Target(MigrationArea.Parse(cfg.SourceRange),cfg.FirstStart,cfg.SecondStart,task.Slot).Address;}catch{}
                    int n=grid.Rows.Add(task.Selected,task.Code,task.Name,task.File,task.Sheet,task.Confirmed,target,task.Status);grid.Rows[n].Tag=task;
                    var cell=new DataGridViewComboBoxCell{DisplayStyle=DataGridViewComboBoxDisplayStyle.DropDownButton,FlatStyle=FlatStyle.Flat};
                    foreach(string name in task.SheetNames.Concat(new[]{task.Sheet.Length>0?task.Sheet:cfg.SourceSheet}).Distinct())cell.Items.Add(name);
                    cell.Value=task.Sheet.Length>0?task.Sheet:cfg.SourceSheet;grid.Rows[n].Cells["sheet"]=cell;
                }
            }
            finally{loading=false;}
        }
        private void Edited(object sender,DataGridViewCellEventArgs e)
        {
            if(loading||e.RowIndex<0||busy)return;var row=grid.Rows[e.RowIndex];var task=row.Tag as NotesLinkTask;if(task==null)return;
            string field=grid.Columns[e.ColumnIndex].Name;
            if(field=="selected")task.Selected=object.Equals(row.Cells[field].Value,true);
            if(field=="confirmed")task.Confirmed=object.Equals(row.Cells[field].Value,true);
            if(field=="file"){task.File=Convert.ToString(row.Cells[field].Value).Trim().Trim('"');task.Sheet="";task.NameNeedsConfirmation=false;task.Confirmed=false;}
            if(field=="sheet"){task.Sheet=Convert.ToString(row.Cells[field].Value);task.NameNeedsConfirmation=false;}
            Changed();
        }
        private void SelectAll(bool value){if(busy)return;foreach(var task in tasks)task.Selected=value;Changed();Render();}
        private NotesLinkTask CurrentTask(){return grid.CurrentRow==null?tasks.FirstOrDefault(t=>t.Selected):(NotesLinkTask)grid.CurrentRow.Tag;}
        private string SampleFile()
        {
            if(tasks.Count==0||leadKey!=Config().LeadKey)ReadCompanies();
            var task=tasks.FirstOrDefault(t=>File.Exists(t.File));if(task!=null)return task.File;
            var files=HostDialogService.OpenFiles(this,"请选择来源工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm");
            if(files.Length==1)return files[0];throw new OperationCanceledException("未选择来源工作簿，原配置保留。");
        }
        private RoundedButton SheetPick(ComboBox field,Func<string> file)
        {
            var button=Ui.Button("选择工作表",130);button.Click+=delegate{Attempt(delegate{
                string path=file();var names=DataMigrationHost.PackageSheets(path);
                using(var dialog=new Form{Text="选择工作表 · "+Path.GetFileName(path),Width=340,Height=350,StartPosition=FormStartPosition.CenterParent})
                {var list=new ListBox{Dock=DockStyle.Fill};list.Items.AddRange(names.Cast<object>().ToArray());var ok=new Button{Text="确定",Dock=DockStyle.Bottom,DialogResult=DialogResult.OK};dialog.Controls.Add(list);dialog.Controls.Add(ok);dialog.AcceptButton=ok;
                 if(dialog.ShowDialog(this)==DialogResult.OK&&list.SelectedItem!=null)field.Text=Convert.ToString(list.SelectedItem);}
            });};return button;
        }
        private string SampleSheet(){return sourceSheet.Text;}
        private void ChangeFile(){var task=CurrentTask();if(task==null)throw new InvalidOperationException("请先选择一家公司。");var files=HostDialogService.OpenFiles(this,"Excel工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm");if(files.Length==1){task.File=files[0];task.Sheet="";task.Confirmed=false;task.NameNeedsConfirmation=false;task.SheetNames=DataMigrationHost.PackageSheets(files[0]);Changed();Render();}}
        private async void Preview()
        {
            if(busy)return;
            try
            {
                if(CanRun!=null&&!CanRun())throw new InvalidOperationException("请等待当前任务结束。");grid.EndEdit();var cfg=Config();if(leadKey!=cfg.LeadKey)throw new InvalidOperationException("父公司、目录或Lead配置已变化，请先重新读取公司。");
                if(DataMigrationHost.Digest(cfg.Parent)!=leadHash)throw new InvalidOperationException("父公司文件在读取名单后已变化，请重新读取公司。");
                var rows=tasks.Select(t=>t.Copy()).ToList();NotesLinkExecution.CheckOpenWorkbooks(Application,rows.Where(t=>t.Selected&&File.Exists(t.File)).Select(t=>t.File),cfg.Parent,false);
                stop=false;plan=null;SetBusy(true);progress.Maximum=Math.Max(1,rows.Count(t=>t.Selected));progress.Value=0;
                var candidate=await Task.Run(()=>NotesLinkPlanner.Build(cfg,rows,delegate(string message){if(!IsDisposed)BeginInvoke((Action)delegate{Append(message);progress.Value=Math.Min(progress.Maximum,progress.Value+1);});},delegate{return stop;}));
                foreach(var task in candidate.Tasks){int index=tasks.FindIndex(t=>t.LeadRow==task.LeadRow);if(index>=0)tasks[index]=task;}
                plan=candidate;Render();int blocked=plan.Tasks.Count(t=>!t.Status.StartsWith("通过",StringComparison.Ordinal));
                state.Text="公司 "+plan.Tasks.Count+" 家；来源 "+candidate.Source.Address+"；覆盖已有内容 "+candidate.Occupied+" 格；阻断 "+blocked+" 项。";
                Append("预检只读，未修改工作簿。结构性合并内部不写入 "+plan.Tasks.Sum(t=>t.Skipped)+" 格，固定槽位不重新排列。");
            }
            catch(Exception error){plan=null;Append(error is OperationCanceledException?error.Message:"预检失败："+error.Message);}
            finally{SetBusy(false);if(closePending)Close();}
        }
        private void Begin()
        {
            if(plan==null||plan.Config.Key!=Config().Key||plan.Tasks.Any(t=>!t.Status.StartsWith("通过",StringComparison.Ordinal)))throw new InvalidOperationException("请重新预检，确保所有勾选任务通过。");
            NotesLinkExecution.CheckOpenWorkbooks(Application,plan.Tasks.Select(t=>t.File),plan.Config.Parent,true);
            string text="将为 "+plan.Tasks.Count+" 家公司生成真实外链，覆盖父公司“"+plan.TargetSheet+"”中 "+plan.Occupied+" 个已有内容单元格。\n\n执行前建立文件备份，全部核验通过后更新父公司原文件；临时工作簿不显示。原有合并、边框和尺寸保留，仅调整金额零值显示。\n\n是否执行？";
            if(UiDialog.Show(this,text,"Notes 批量链接",MessageBoxButtons.YesNo,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2)!=DialogResult.Yes)return;
            stop=false;execution=new NotesLinkExecution(Application,plan,Append);sequence=execution.Run().GetEnumerator();progress.Maximum=plan.Tasks.Count;progress.Value=0;SetBusy(true);timer.Start();
        }
        private void Tick()
        {
            if(ticking||sequence==null)return;ticking=true;timer.Stop();
            try
            {
                if(sequence.MoveNext()){state.Text=sequence.Current;Append(sequence.Current);progress.Value=Math.Min(progress.Maximum,execution.Completed);return;}
                lastParent=plan.Config.Parent;lastBackup=execution.Backup;lastHash=execution.CommittedHash;Append("全部交付完成。备份："+lastBackup);Finish(null);
            }
            catch(Exception error){Finish(error);}
            finally{ticking=false;if(sequence!=null)timer.Start();}
        }
        private void Finish(Exception error)
        {
            bool committed=execution!=null&&execution.Committed;
            if(committed){lastParent=plan.Config.Parent;lastBackup=execution.Backup;lastHash=execution.CommittedHash;}
            timer.Stop();if(sequence!=null){sequence.Dispose();sequence=null;}if(execution!=null){execution.Dispose();execution=null;}
            if(error!=null){foreach(var task in plan.Tasks)task.Status=committed?"已更新／后处理警告":error is OperationCanceledException?"已取消／未交付":"中止／请检查恢复记录";Append((committed?"父公司已更新，后处理警告：":"本批中止，详细恢复路径见日志：")+error.Message);}
            Render();plan=null;SetBusy(false);if(closePending)Close();
        }
        private void Restore()
        {
            if(string.IsNullOrEmpty(lastBackup))throw new InvalidOperationException("本窗口没有可还原的更新。");NotesLinkExecution.CheckOpenWorkbooks(Application,new string[0],lastParent,true);
            if(UiDialog.Show(this,"将用备份还原父公司；当前版本也会保留。如果更新后又被修改，将阻止自动还原。是否继续？","还原Notes更新",MessageBoxButtons.YesNo,MessageBoxIcon.Warning)!=DialogResult.Yes)return;
            string displaced=MigrationOutputTransaction.Restore(lastParent,lastBackup,lastHash);Append("已还原；还原前版本保留："+displaced);lastBackup=null;undo.Enabled=false;Changed();
        }
        private void Parameters()
        {
            if(loading||busy)return;
            try{int first=MigrationPlanner.ColumnIndex(firstColumn.Text),last=MigrationPlanner.ColumnIndex(lastColumn.Text);if(last<first)return;sourceRange.Text=new MigrationArea{Row=(int)sourceFirstRow.Value,Column=first,Rows=(int)sourceRows.Value,Columns=last-first+1}.Address;}catch{}
        }
        private void SyncRange()
        {
            var area=MigrationArea.Parse(sourceRange.Text);loading=true;
            try{firstColumn.Text=CellAddress.ColumnLetters(area.Column);lastColumn.Text=CellAddress.ColumnLetters(area.Column+area.Columns-1);sourceFirstRow.Value=area.Row;sourceRows.Value=area.Rows;}finally{loading=false;}
        }
        private void TargetAnchor(MigrationArea area)
        {
            if(area.Rows==1&&area.Columns==1)return;
            string sheet=NotesLinkRules.Resolve(DataMigrationHost.PackageSheets(parent.Text),targetSheet.Text);
            if(!MigrationSheetSnapshot.Load(parent.Text,sheet).Merges.Any(m=>NotesLinkRules.Same(m,area)))throw new InvalidOperationException("起点请选择一个单元格或一个完整合并单元格。");
        }
        private static void Single(MigrationArea area){if(area.Rows!=1||area.Columns!=1)throw new InvalidOperationException("请选择一个单元格。");}
        private RoundedButton ColumnPick(TextBox field){return Pick("Lead字段",delegate{return parent.Text;},delegate{return leadSheet.Text;},delegate(MigrationArea area){if(area.Columns!=1)throw new InvalidOperationException("请选择单列中的单元格或区域。");field.Text=CellAddress.ColumnLetters(area.Column);});}
        private RoundedButton Pick(string role,Func<string> file,Func<string> sheet,Action<MigrationArea> apply)
        {
            var button=Ui.Button("在Excel中选择",130);button.Click+=delegate{Attempt(delegate{string path=file(),name=NotesLinkRules.Resolve(DataMigrationHost.PackageSheets(path),sheet());string address=null;bool previousTop=TopMost;TopMost=false;Hide();try{using(var host=new DataMigrationHost(Application))address=host.PickSourceRange(path,name,"A1",null,role,delegate(Action restore){MigrationForeground.Bring(MigrationForeground.ActiveFrame(Application,IntPtr.Zero));});}finally{TopMost=previousTop;Show();Activate();}if(address!=null)apply(MigrationArea.Parse(address));});};return button;
        }
        private RoundedButton Current(string role,Func<string> file,Func<string> expectedSheet,Action<MigrationArea> apply)
        {
            var button=Ui.Button("读取当前选区",125);button.Click+=delegate{Attempt(delegate{
                using(var bag=new ComBag()){
                    dynamic selection=bag.Own(((dynamic)Application).Selection);if(selection==null)throw new InvalidOperationException("请在Excel中选择单元格区域。");dynamic sheet=bag.Own(selection.Worksheet),book=bag.Own(sheet.Parent);
                    string path=file(),name=NotesLinkRules.Resolve(DataMigrationHost.PackageSheets(path),expectedSheet());if(!string.Equals(Convert.ToString(book.FullName),Path.GetFullPath(path),StringComparison.OrdinalIgnoreCase)||Convert.ToString(sheet.Name)!=name)throw new InvalidOperationException("请选择指定工作簿的“"+name+"”表中的"+role+"区域。");
                    apply(DataMigrationHost.ReadCoordinates((object)selection,(object)book,(object)sheet));
                }});};return button;
        }
        internal void Shutdown(){if(busy){closePending=true;stop=true;if(execution!=null)execution.Cancelled=true;}else Close();}
    }
}
