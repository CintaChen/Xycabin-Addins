using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Excel_Iterate
{
    // STA iterator: UI gets control between bounded formula blocks. The parent
    // on disk changes only after close/reopen verification and atomic commit.
    internal sealed class NotesLinkExecution : IDisposable
    {
        private readonly NotesLinkPlan plan;
        private readonly object sourceApp;
        private readonly Action<string> log;
        private readonly ComBag references=new ComBag();
        private HiddenWorkbookApplication hidden;
        private object book,sheet;
        private string stage,batch,manifest;
        private bool closed=true,disposed;
        internal bool Cancelled,Committed;
        internal string Backup="",PreBackup="",CommittedHash="";
        internal int Completed,FormulaWrites,ValueErrors,LinkWarnings;
        internal Func<int,bool> FailureForProbe;
        internal NotesLinkExecution(object application,NotesLinkPlan frozen,Action<string> output)
        {sourceApp=application;plan=frozen;log=output??delegate{};}
        internal static void CheckOpenWorkbooks(object app,IEnumerable<string> sources,string parent,bool requireClosed)
        {
            var relevant=new HashSet<string>(sources.Select(Path.GetFullPath),StringComparer.OrdinalIgnoreCase);relevant.Add(Path.GetFullPath(parent));
            using(var bag=new ComBag())
            {
                dynamic books=bag.Own(((dynamic)app).Workbooks);
                for(int i=1;i<=Convert.ToInt32(books.Count);i++)using(var one=new ComBag())
                {
                    dynamic current=one.Own(books.Item(i));string file=Convert.ToString(current.FullName);if(!Path.IsPathRooted(file)||!relevant.Contains(Path.GetFullPath(file)))continue;
                    if(!Convert.ToBoolean(current.Saved))throw new InvalidOperationException("请先手动保存工作簿的修改，再预检："+file);
                    if(requireClosed&&string.Equals(Path.GetFullPath(file),Path.GetFullPath(parent),StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("备份后更新原文件需要父公司工作簿处于关闭状态；请手动关闭后重试，不会自动保存或关闭它。");
                }
            }
        }
        private void Check()
        {
            if(Cancelled)throw new OperationCanceledException("已取消；父公司原文件未交付修改。");
        }
        private void CheckParent(){Check();if(DataMigrationHost.Digest(plan.Config.Parent)!=plan.ParentHash)throw new InvalidOperationException("父公司文件在预检后变化，请重新预检。");}
        private void Record(string state)
        {
            if(manifest==null)return;
            var values=MigrationJson.Obj("state",state,"parent",plan.Config.Parent,"originalHash",plan.ParentHash,"stage",stage,"preBackup",PreBackup,"backup",Backup,"committedHash",CommittedHash);
            File.WriteAllText(manifest,MigrationJson.Write(values),new System.Text.UTF8Encoding(false));
        }
        private void Open(bool readOnly)
        {
            using(var one=new ComBag())
            {dynamic books=one.Own(((dynamic)hidden.Application).Workbooks);book=references.Own(books.Open(stage,0,readOnly));if(book==null||(!readOnly&&Convert.ToBoolean(((dynamic)book).ReadOnly)))throw new InvalidOperationException("无法打开Notes工作副本。");closed=false;dynamic sheets=one.Own(((dynamic)book).Worksheets);sheet=references.Own(sheets.Item(plan.TargetSheet));}
        }
        private void Close()
        {if(book!=null&&!closed){((dynamic)book).Close(false);closed=true;book=null;sheet=null;}}
        private IEnumerable<MigrationArea> Chunks(MigrationArea area)
        {for(int r=0;r<area.Rows;r+=128)yield return new MigrationArea{Row=area.Row+r,Column=area.Column,Rows=Math.Min(128,area.Rows-r),Columns=area.Columns};}
        private object[,] Matrix(NotesLinkTask task,MigrationArea area)
        {
            return DataMigrationEngine.LinkMatrix(task.File,task.Sheet,plan.Source.Row+area.Row-task.Target.Row,plan.Source.Column+area.Column-task.Target.Column,area.Rows,area.Columns);
        }
        private void Format(MigrationArea area)
        {
            using(var bag=new ComBag())
            {
                dynamic range=bag.Own(((dynamic)sheet).Range[area.Address]);object uniform=range.NumberFormat;
                if(uniform!=null&&uniform!=DBNull.Value){range.NumberFormat=NotesLinkRules.ZeroFormat(Convert.ToString(uniform));return;}
            }
            // Preserve border/font/alignment styles. Group format operations by
            // original package style instead of fetching every cell through COM.
            var groups=new Dictionary<int,List<MigrationArea>>();
            for(int r=area.Row;r<area.Row+area.Rows;r++)for(int c=area.Column;c<area.Column+area.Columns;)
            {
                int style=plan.Original.Style(r,c),start=c++;while(c<area.Column+area.Columns&&plan.Original.Style(r,c)==style)c++;
                List<MigrationArea> list;if(!groups.TryGetValue(style,out list))groups[style]=list=new List<MigrationArea>();list.Add(new MigrationArea{Row=r,Column=start,Rows=1,Columns=c-start});
            }
            foreach(var group in groups.Values)
            {
                string expected=null;
                var chunks=new List<string>();string joined="";
                foreach(var run in group){string item=run.Address;if(joined.Length+item.Length+1>200){chunks.Add(joined);joined="";}joined+=joined.Length==0?item:","+item;}if(joined.Length>0)chunks.Add(joined);
                foreach(var address in chunks)using(var bag=new ComBag())
                {
                    dynamic range=bag.Own(((dynamic)sheet).Range[address]);object format=range.NumberFormat;
                    if(format==null||format==DBNull.Value)throw new InvalidOperationException("无法安全处理混合数字格式："+address);
                    string actual=Convert.ToString(format);if(expected==null)expected=actual;else if(actual!=expected)throw new InvalidOperationException("目标有效样式与预检不一致："+address);
                }
                string desired=NotesLinkRules.ZeroFormat(expected);
                foreach(var address in chunks)using(var bag=new ComBag()){dynamic range=bag.Own(((dynamic)sheet).Range[address]);range.NumberFormat=desired;}
            }
        }
        private void Verify(NotesLinkTask task,MigrationArea area,bool values)
        {
            using(var bag=new ComBag())
            {
                dynamic range=bag.Own(((dynamic)sheet).Range[area.Address]);object actual=range.Formula;var expected=Matrix(task,area);
                for(int r=0;r<area.Rows;r++)for(int c=0;c<area.Columns;c++)
                {
                    if(!string.Equals(Convert.ToString(SummaryArrays.Pick(actual,r,area.Rows,c,area.Columns)),Convert.ToString(expected[r,c]),StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("公式读回不一致："+CellAddress.Name(area.Row+r,area.Column+c));
                }
                object formula=range.HasFormula;if(formula==null||formula==DBNull.Value||!Convert.ToBoolean(formula))throw new InvalidOperationException("存在未写成真实公式的单元格："+area.Address);
                if(values)
                {
                    object data=range.Value2;var errors=new System.Text.StringBuilder();
                    for(int r=0;r<area.Rows;r++)for(int c=0;c<area.Columns;c++)
                    {object value=SummaryArrays.Pick(data,r,area.Rows,c,area.Columns);if(SummaryAreaEngine.IsError(value)){ValueErrors++;errors.AppendLine("数值异常｜"+task.Name+"｜"+plan.TargetSheet+"!"+CellAddress.Name(area.Row+r,area.Column+c)+"｜"+NotesLinkRules.ErrorText(value));}}
                    if(errors.Length>0)log(errors.ToString().TrimEnd());
                }
            }
        }
        private void RestoreGeometry()
        {
            var rows=new SortedDictionary<int,double>();
            foreach(var task in plan.Tasks)
            {for(int r=task.Target.Row;r<task.Target.Row+task.Target.Rows;r++)rows[r]=plan.Original.Height(r);}
            RestoreRows(rows);
            foreach(int row in plan.Original.HiddenRows.Where(rows.ContainsKey))using(var bag=new ComBag()){dynamic range=bag.Own(((dynamic)sheet).Range[row+":"+row]);range.Hidden=true;}
        }
        private void RestoreRows(SortedDictionary<int,double> values)
        {
            var all=values.ToList();for(int i=0;i<all.Count;)
            {
                int first=all[i].Key,last=first;double size=all[i].Value;i++;
                while(i<all.Count&&all[i].Key==last+1&&Math.Abs(all[i].Value-size)<0.001){last=all[i].Key;i++;}
                using(var bag=new ComBag())
                {dynamic range=bag.Own(((dynamic)sheet).Range[first+":"+last]);range.RowHeight=size;}
            }
        }
        internal IEnumerable<string> Run()
        {
            if(plan.Tasks.Count==0||plan.Tasks.Any(t=>!t.Status.StartsWith("通过",StringComparison.Ordinal)))throw new InvalidOperationException("请先让所有勾选任务通过预检。");
            CheckOpenWorkbooks(sourceApp,plan.Tasks.Select(t=>t.File),plan.Config.Parent,true);CheckParent();
            try{using(File.Open(plan.Config.Parent,FileMode.Open,FileAccess.ReadWrite,FileShare.None)){}}
            catch(IOException error){throw new IOException("父公司文件被其他Excel进程或同步程序占用，请关闭或等待完成后重试；未覆盖原文件。",error);}
            batch="Notes_"+DateTime.Now.ToString("yyyyMMdd_HHmmss",CultureInfo.InvariantCulture)+"_"+Guid.NewGuid().ToString("N").Substring(0,8);
            string directory=Path.Combine(Path.GetDirectoryName(plan.Config.Parent),"复制备份",batch);Directory.CreateDirectory(directory);
            PreBackup=Path.Combine(directory,"执行前_"+Path.GetFileName(plan.Config.Parent));SummaryFileBatch.CopyInputSnapshot(plan.Config.Parent,PreBackup);
            if(DataMigrationHost.Digest(PreBackup)!=plan.ParentHash)throw new InvalidOperationException("执行前备份核验失败，未写入父公司。");
            stage=Path.Combine(Path.GetDirectoryName(plan.Config.Parent),".xy-notes-"+Guid.NewGuid().ToString("N")+Path.GetExtension(plan.Config.Parent));manifest=Path.Combine(directory,"session.json");
            File.Copy(PreBackup,stage,false);File.SetAttributes(stage,FileAttributes.Normal);Record("Prepared");yield return "已建立执行前备份："+PreBackup;
            Check();hidden=new HiddenWorkbookApplication(sourceApp);Open(false);Record("Writing");
            plan.Original.StandardHeight=Convert.ToDouble(((dynamic)sheet).StandardHeight);plan.Original.StandardWidth=Convert.ToDouble(((dynamic)sheet).StandardWidth);
            foreach(var task in plan.Tasks)
            {
                Check();if(DataMigrationHost.Digest(task.File)!=task.Hash)throw new InvalidOperationException("来源文件在预检后变化："+task.File);
                task.Status="执行中";yield return "开始："+task.Code+task.Name+" → "+plan.TargetSheet+"!"+task.Target.Address;
                foreach(var part in task.Blocks.Concat(task.Anchors))foreach(var chunk in Chunks(part))
                {
                    Check();Format(chunk);
                    using(var bag=new ComBag()){dynamic range=bag.Own(((dynamic)sheet).Range[chunk.Address]);DataMigrationEngine.AssignFormula((object)range,Matrix(task,chunk));}
                    Verify(task,chunk,false);FormulaWrites+=chunk.Rows*chunk.Columns;yield return "写入："+task.Name+"｜"+chunk.Address+"｜累计公式 "+FormulaWrites;
                    if(FailureForProbe!=null&&FailureForProbe(FormulaWrites))throw new InvalidOperationException("测试注入的写入异常。");
                }
                Completed++;task.Status="公式已写入";
            }
            foreach(var task in plan.Tasks)
            {
                Check();yield return "更新链接："+task.Name;
                try{((dynamic)book).UpdateLink(task.File,1);log("链接更新完成："+task.Name);}catch(Exception error){LinkWarnings++;log("链接更新警告："+task.Name+"｜"+error.Message);}
            }
            Check();yield return "正在重算并保存临时副本…";
            try{((dynamic)sheet).Calculate();}catch(Exception error){LinkWarnings++;log("计算警告（保留真实公式）："+error.Message);}
            RestoreGeometry();var macros=DataMigrationHost.MacroFingerprints(plan.Config.Parent);((dynamic)book).Save();Close();
            CheckParent();
            if(!DataMigrationHost.PackageSheets(stage).SequenceEqual(DataMigrationHost.PackageSheets(plan.Config.Parent)))throw new InvalidOperationException("副本工作表名称或顺序发生变化，未交付。");
            var afterMacros=DataMigrationHost.MacroFingerprints(stage);if(!macros.OrderBy(x=>x.Key).SequenceEqual(afterMacros.OrderBy(x=>x.Key)))throw new InvalidOperationException("VBA内容发生变化，未交付。");
            Record("Saved");yield return "重开副本核验公式、合并结构及样式…";Open(true);
            var facts=NotesSheetFacts.Load(stage,plan.TargetSheet);facts.StandardWidth=Convert.ToDouble(((dynamic)sheet).StandardWidth);facts.StandardHeight=Convert.ToDouble(((dynamic)sheet).StandardHeight);facts.VerifyAgainst(plan.Original,plan.Tasks);
            var links=((dynamic)book).LinkSources(1) as Array;
            if(links!=null&&links.Cast<object>().Any(x=>Convert.ToString(x).IndexOf(Path.GetFileName(stage),StringComparison.OrdinalIgnoreCase)>=0))throw new InvalidOperationException("检测到指向临时副本的外链，未交付。");
            foreach(var task in plan.Tasks)foreach(var part in task.Blocks.Concat(task.Anchors))foreach(var chunk in Chunks(part)){Check();Verify(task,chunk,true);yield return "核验："+task.Name+"｜"+chunk.Address;}
            Close();foreach(var task in plan.Tasks)if(DataMigrationHost.Digest(task.File)!=task.Hash)throw new InvalidOperationException("来源文件在执行期间变化，未更新父公司："+task.File);
            CheckParent();Record("Verified");yield return "核验完成，即将备份后更新父公司原文件…";CheckParent();
            string verifiedHash=DataMigrationHost.Digest(stage);
            try{Backup=MigrationOutputTransaction.Commit(stage,plan.Config.Parent,plan.ParentHash,batch);}
            catch(Exception error){log("提交中止，请检查父公司及恢复记录："+manifest+"｜"+error.Message);throw;}
            Committed=true;CommittedHash=verifiedHash;
            try{Record("Committed");}catch(Exception error){log("父公司已更新，但提交记录写入失败；恢复备份："+Backup+"｜"+error.Message);}
            foreach(var task in plan.Tasks)task.Status=LinkWarnings>0||ValueErrors>0?"已更新（有链接／数值警告）":"已更新";
            yield return "已更新父公司｜公式 "+FormulaWrites+"；链接警告 "+LinkWarnings+"；数值异常 "+ValueErrors+"；恢复备份："+Backup;
        }
        public void Dispose()
        {
            if(disposed)return;disposed=true;
            try{Close();}catch(Exception error){log("临时工作簿关闭失败："+error.Message);}
            try{references.Dispose();}catch(Exception error){log("COM释放失败："+error.Message);}
            try{if(hidden!=null)hidden.Dispose();}catch(Exception error){log("后台进程退出失败："+error.Message);}hidden=null;
            if(!Committed)
            {
                try{Record(Cancelled?"Cancelled":"Failed");}catch(Exception error){log("恢复记录写入失败："+error.Message);}
                try{if(stage!=null&&File.Exists(stage))File.Delete(stage);}catch(Exception error){log("临时副本保留待清理："+stage+"｜"+error.Message);}
                if(PreBackup.Length>0)log("本批未交付；执行前备份保留："+PreBackup);
            }
        }
    }
}
