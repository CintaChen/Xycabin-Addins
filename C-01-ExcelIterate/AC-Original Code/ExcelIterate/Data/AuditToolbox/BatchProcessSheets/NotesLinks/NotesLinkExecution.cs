using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Excel_Iterate
{
    // STA iterator: UI gets control between bounded formula blocks. The parent
    // on disk changes only after saving the private copy and atomic commit.
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
        {for(int r=0;r<area.Rows;r+=512)yield return new MigrationArea{Row=area.Row+r,Column=area.Column,Rows=Math.Min(512,area.Rows-r),Columns=area.Columns};}
        private object[,] Matrix(NotesLinkTask task,MigrationArea area)
        {
            return DataMigrationEngine.LinkMatrix(task.File,task.Sheet,plan.Source.Row+area.Row-task.Target.Row,plan.Source.Column+area.Column-task.Target.Column,area.Rows,area.Columns);
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
            Check();hidden=new HiddenWorkbookApplication(sourceApp);Open(false);
            dynamic worker=hidden.Application;worker.Calculation=-4135;worker.CalculateBeforeSave=false;Record("Writing");
            plan.Original.StandardHeight=Convert.ToDouble(((dynamic)sheet).StandardHeight);plan.Original.StandardWidth=Convert.ToDouble(((dynamic)sheet).StandardWidth);
            foreach(var task in plan.Tasks)
            {
                Check();if(DataMigrationHost.Digest(task.File)!=task.Hash)throw new InvalidOperationException("来源文件在预检后变化："+task.File);
                task.Status="执行中";yield return "开始："+task.Code+task.Name+" → "+plan.TargetSheet+"!"+task.Target.Address;
                var refresh=System.Diagnostics.Stopwatch.StartNew();
                foreach(var part in task.Blocks.Concat(task.Anchors))foreach(var chunk in Chunks(part))
                {
                    Check();
                    using(var bag=new ComBag()){dynamic range=bag.Own(((dynamic)sheet).Range[chunk.Address]);DataMigrationEngine.AssignFormula((object)range,Matrix(task,chunk));}
                    FormulaWrites+=chunk.Rows*chunk.Columns;
                    if(refresh.ElapsedMilliseconds>=200){yield return "粘贴链接："+task.Name+"｜累计 "+FormulaWrites+" 格";refresh.Restart();}
                    if(FailureForProbe!=null&&FailureForProbe(FormulaWrites))throw new InvalidOperationException("测试注入的写入异常。");
                }
                Completed++;task.Status="链接已粘贴";yield return "已粘贴："+task.Name+"｜累计 "+FormulaWrites+" 格";
            }
            Check();yield return "正在保存链接，不更新外链或重新计算…";
            RestoreGeometry();((dynamic)book).Save();Close();Record("Saved");
            foreach(var task in plan.Tasks)if(DataMigrationHost.Digest(task.File)!=task.Hash)throw new InvalidOperationException("来源文件在执行期间变化，未更新汇总文件："+task.File);
            CheckParent();yield return "链接已保存，即将更新汇总文件…";CheckParent();
            string verifiedHash=DataMigrationHost.Digest(stage);
            try{Backup=MigrationOutputTransaction.Commit(stage,plan.Config.Parent,plan.ParentHash,batch);}
            catch(Exception error){log("提交中止，请检查父公司及恢复记录："+manifest+"｜"+error.Message);throw;}
            Committed=true;CommittedHash=verifiedHash;
            try{Record("Committed");}catch(Exception error){log("父公司已更新，但提交记录写入失败；恢复备份："+Backup+"｜"+error.Message);}
            foreach(var task in plan.Tasks)task.Status="链接已粘贴";
            yield return "已更新汇总文件｜粘贴链接 "+FormulaWrites+" 格；未更新外链或检查数值；恢复备份："+Backup;
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
