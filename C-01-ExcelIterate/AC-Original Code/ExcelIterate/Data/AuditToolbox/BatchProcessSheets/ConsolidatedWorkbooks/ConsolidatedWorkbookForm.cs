using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class ConsolidatedWorkbookForm : Form
    {
        private readonly TextBox template=new TextBox{ReadOnly=true,Dock=DockStyle.Fill},output=new TextBox{Dock=DockStyle.Fill};
        internal readonly CompanyStructurePicker Structure=new CompanyStructurePicker{SimpleMode=true,SelectParents=true,Dock=DockStyle.Fill};
        private readonly RichTextBox log=new RichTextBox{Dock=DockStyle.Fill,ReadOnly=true,BackColor=Color.White};
        private readonly ProgressBar progress=new ProgressBar{Dock=DockStyle.Fill,Style=ProgressBarStyle.Continuous};
        private readonly Label state=new Label{AutoSize=true,Text="等待预检"};
        private readonly TableLayoutPanel settings=new TableLayoutPanel{Dock=DockStyle.Top,AutoSize=true,ColumnCount=4};
        private readonly RoundedButton preview=Ui.Button("读取并预检",130),start=Ui.Button("开始生成",120),cancel=Ui.Button("取消任务",100);
        private readonly Timer timer=new Timer{Interval=30};
        private CompanyWorkbookHost host;
        private List<ConsolidatedWorkbookPlan> plans;
        private int position,success,failed;
        private bool ticking,closePending,cancelled,preparing;
        internal Func<string,CompanyWorkbookHost> OpenTemplate;
        internal Func<CompanyWorkbookHost> OpenCurrentTemplate;
        internal Func<bool> CanRun;
        internal bool RememberConfigurationEnabled;
        internal event Action<bool> BusyChanged;
        private bool Running{get{return plans!=null;}}
        internal ConsolidatedWorkbookForm()
        {
            Text="合并模板生成";DialogSizing.Prepare(this,"company.consolidated","table",Ui.Blue);
            var root=new TableLayoutPanel{Dock=DockStyle.Fill,Padding=new Padding(16),ColumnCount=1,RowCount=6};
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));root.RowStyles.Add(new RowStyle(SizeType.Percent,65));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));root.RowStyles.Add(new RowStyle(SizeType.Absolute,20));root.RowStyles.Add(new RowStyle(SizeType.Percent,35));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var browse=Ui.Button("选择模板",110);var current=Ui.Button("使用当前工作簿",150);var folder=Ui.Button("选择文件夹",110);
            AddSetting(0,"合并模板",template,browse,current);AddSetting(1,"输出文件夹",output,folder,null);
            root.Controls.Add(settings,0,0);root.Controls.Add(Structure,0,1);root.Controls.Add(state,0,2);root.Controls.Add(progress,0,3);root.Controls.Add(log,0,4);
            var actions=Ui.ActionBar();var close=Ui.Button("关闭",90);actions.Controls.AddRange(new Control[]{preview,start,cancel,close});root.Controls.Add(actions,0,5);Controls.Add(root);
            browse.Click+=delegate{Attempt(delegate{var paths=HostDialogService.OpenFiles(this,"Excel模板 (*.xlsx)|*.xlsx");if(paths.Length==1)LoadTemplate(paths[0]);});};
            current.Click+=delegate{Attempt(delegate{LoadTemplate(null);});};folder.Click+=delegate{string path=HostDialogService.Folder(this,"选择合并模板输出文件夹",true);if(!string.IsNullOrEmpty(path))output.Text=path;};
            Structure.Log=Append;Structure.Changed+=delegate{if(!Running)state.Text="配置已改变，请重新预检";};output.TextChanged+=delegate{if(!Running)state.Text="输出目录已改变，请重新预检";};
            preview.Click+=delegate{Attempt(delegate{Check();});};start.Click+=delegate{Attempt(Begin);};cancel.Click+=delegate{cancelled=true;Append("取消已请求，将在当前文件完成后停止。");};close.Click+=delegate{Close();};cancel.Enabled=false;
            timer.Tick+=delegate{Tick();};
            FormClosing+=delegate(object sender,FormClosingEventArgs e){if(Running||ticking||preparing){e.Cancel=true;closePending=true;cancelled=true;}else Remember();};
            Disposed+=delegate{timer.Stop();timer.Dispose();if(host!=null){host.Dispose();host=null;}};
        }
        private void AddSetting(int row,string title,Control field,Control button,Control extra)
        {settings.RowStyles.Add(new RowStyle(SizeType.AutoSize));settings.Controls.Add(new Label{Text=title,AutoSize=true,Anchor=AnchorStyles.Left,Margin=new Padding(3,8,10,8)},0,row);settings.Controls.Add(field,1,row);settings.Controls.Add(button,2,row);if(extra!=null)settings.Controls.Add(extra,3,row);}
        private void Append(string message){log.AppendText(DateTime.Now.ToString("HH:mm:ss")+"  "+message+Environment.NewLine);log.SelectionStart=log.TextLength;log.ScrollToCaret();log.Update();state.Update();}
        private void Attempt(Action action){try{if(Running||preparing||CanRun!=null&&!CanRun())throw new InvalidOperationException("请等待当前任务完成。");action();}catch(Exception error){Append("失败："+error.Message);}}
        private void Busy(bool value){settings.Enabled=!value;Structure.Enabled=!value;preview.Enabled=start.Enabled=!value;cancel.Enabled=value;if(BusyChanged!=null)BusyChanged(value);}
        private void LoadTemplate(string path)
        {
            CompanyWorkbookHost candidate=null;preparing=true;Busy(true);
            try
            {
                if(path==null){if(OpenCurrentTemplate==null)throw new InvalidOperationException("模板服务未连接。");candidate=OpenCurrentTemplate();}
                else{if(OpenTemplate==null)throw new InvalidOperationException("模板服务未连接。");candidate=OpenTemplate(path);}
                var profile=ConsolidatedTemplateProfile.Read(candidate.Snapshot);
                if(host!=null)host.Dispose();host=candidate;candidate=null;template.Text=path??host.TemplateSource;
                Append("模板已读取。仅写入Lead的K:O公司名单；本部在前，随后为全部直系子公司。模板公式公司范围截止第"+profile.CapacityLast+"行。");
            }
            finally{if(candidate!=null)candidate.Dispose();preparing=false;Busy(false);}
        }
        private List<ConsolidatedWorkbookPlan> Check()
        {
            if(host==null)throw new InvalidOperationException("请先选择合并模板。");
            var all=Structure.All();var selected=Structure.Selected();if(selected.Count==0)throw new InvalidOperationException("请勾选需要生成的合并主体。");
            CompanyStructureTasks.Classify(all,true);var profile=ConsolidatedTemplateProfile.Read(host.Snapshot);var result=new List<ConsolidatedWorkbookPlan>();
            var duplicates=new HashSet<string>(selected.GroupBy(t=>t.FileName,StringComparer.OrdinalIgnoreCase).Where(g=>g.Count()>1).Select(g=>g.Key),StringComparer.OrdinalIgnoreCase);
            foreach(var parent in selected)
            {
                try
                {
                    if(duplicates.Contains(parent.FileName))throw new InvalidOperationException("勾选任务使用了相同的输出文件名。");
                    var plan=profile.Plan(all,parent,output.Text);result.Add(plan);Structure.Report(parent.Row,"可生成："+(plan.Entries.Count-1)+"家（含本部）");
                    Append(plan.Record.Name+" → Lead K2:O"+plan.ClearLast+"；名单："+string.Join("、",plan.Entries.Select(e=>e.Code+e.Name)));
                }
                catch(Exception error){Structure.Report(parent.Row,"阻止："+error.Message);Append("阻止该主体「"+parent.FileName+"」："+error.Message);}
            }
            state.Text="预检通过 "+result.Count+" / "+selected.Count+" 项；尚未生成文件";return result;
        }
        private void Begin()
        {
            var ready=Check();if(ready.Count==0)throw new InvalidOperationException("没有通过预检的合并主体。");
            if(UiDialog.Show(this,"将新建"+ready.Count+"个合并工作簿，并按日志中的范围重写副本Lead公司名单。原模板不变，同名文件不覆盖。是否继续？","合并模板生成",MessageBoxButtons.OKCancel,MessageBoxIcon.Question)!=DialogResult.OK)return;
            Remember();plans=ready;position=success=failed=0;cancelled=false;progress.Maximum=plans.Count;progress.Value=0;Busy(true);timer.Start();
        }
        private void Tick()
        {
            if(ticking||!Running)return;if(cancelled||position>=plans.Count){Finish();return;}
            ticking=true;var plan=plans[position];state.Text="正在生成 "+(position+1)+" / "+plans.Count+"："+plan.Record.Name;
            try
            {
                var request=new CompanyWorkbookRequest{Output=output.Text,UpdateLead=true,TargetSheet="Lead"};
                host.Generate(request,plan.Record,delegate(object book,CompanyWorkbookRequest r,CompanyWorkbookRecord item){ConsolidatedLeadWriter.Write(book,plan);},Append,delegate(object book,CompanyWorkbookRequest r,CompanyWorkbookRecord item){ConsolidatedLeadWriter.Verify(book,plan);});
                success++;Structure.Report(plan.Record.Row,"已生成");Append("完成："+plan.Record.OutputPath);
            }
            catch(Exception error){failed++;Structure.Report(plan.Record.Row,"生成失败");Append("生成失败："+plan.Record.Name+"；"+error.Message);}
            finally{position++;progress.Value=position;ticking=false;state.Text="已完成 "+position+" / "+plans.Count+"；成功 "+success+"，失败 "+failed;}
        }
        private void Finish(){timer.Stop();plans=null;Busy(false);Append(cancelled?"已取消剩余任务，已成功生成的文件保留。":"本轮生成结束。");if(closePending)Close();}
        internal void Shutdown(){if(Running||ticking||preparing){closePending=true;cancelled=true;}else Close();}
        private static string LastFile{get{return Path.Combine(Path.GetDirectoryName(BatchPresetStore.Library("company")),"consolidated-last-settings.json");}}
        private void Remember(){if(!RememberConfigurationEnabled||host==null||!Directory.Exists(output.Text)||!Path.IsPathRooted(template.Text))return;try{BatchPresetStore.Write(LastFile,new BatchPreset{Feature="company",Name="合并模板生成上次设置",Context=template.Text,Values=new Dictionary<string,string>{{"output",output.Text}}},true);}catch(Exception error){Append("设置未保存："+error.Message);}}
        internal void RestoreLastConfiguration(){if(!RememberConfigurationEnabled||!File.Exists(LastFile))return;try{var preset=BatchPresetStore.Read(LastFile,"company");string value;if(preset.Values.TryGetValue("output",out value))output.Text=value;if(File.Exists(preset.Context))LoadTemplate(preset.Context);}catch(Exception error){Append("请重新选择模板："+error.Message);}}
    }
}
