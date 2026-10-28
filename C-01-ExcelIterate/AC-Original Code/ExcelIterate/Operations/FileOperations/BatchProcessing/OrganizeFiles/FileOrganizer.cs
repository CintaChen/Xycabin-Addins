using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Excel_Iterate
{
    internal sealed class OrganizeOptions
    {
        internal bool Replace, Classify, Filter;
        internal HashSet<string> Types = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    }

    internal sealed class OrganizeGroup
    {
        internal string Target; internal FileSnapshot Winner;
        internal List<FileSnapshot> Backups = new List<FileSnapshot>();
    }

    internal sealed class OrganizePlan
    {
        internal string Root, BackupRoot;
        internal List<OrganizeGroup> Groups = new List<OrganizeGroup>();
        internal List<FileSnapshot> Recycle = new List<FileSnapshot>();
        internal List<string> Notes = new List<string>();
        internal List<BatchDirectory> OriginalDirectories=new List<BatchDirectory>();
        internal List<string> ProtectedDirectories=new List<string>();
        internal string Confirmation
        {
            get
            {
                var lines=new StringBuilder();
                lines.AppendLine("移动"+Groups.Count(g=>!string.Equals(g.Winner.Path,g.Target,StringComparison.OrdinalIgnoreCase))+"项；同名备份"+Groups.Sum(g=>g.Backups.Count)+"项；未选类型暂存"+Recycle.Count+"项。");
                lines.AppendLine("确认操作将回收以下原目录及剩余内容：");
                if(OriginalDirectories.Count==0)lines.AppendLine("无原目录需要回收");
                else foreach(var dir in OriginalDirectories)lines.AppendLine(dir.Path);
                return lines.ToString().TrimEnd('\r','\n');
            }
        }
        internal string Preview
        {
            get
            {
                var lines = new StringBuilder();
                lines.AppendLine("移动 " + Groups.Count(g => !string.Equals(g.Winner.Path, g.Target, StringComparison.OrdinalIgnoreCase)) + " 项；同名备份 " + Groups.Sum(g => g.Backups.Count) + " 项；未选类型暂存 " + Recycle.Count + " 项。");
                foreach (var group in Groups) { lines.AppendLine("保留/移动：" + group.Winner.Path + " → " + group.Target); foreach (var item in group.Backups) lines.AppendLine("同名备份：" + item.Path); }
                foreach (var item in Recycle) lines.AppendLine("暂存到备份：" + item.Path);
                foreach (string note in Notes) lines.AppendLine(note);
                if (Groups.Any(g => g.Backups.Count > 0)) lines.AppendLine("备份目录：" + BackupRoot);
                lines.AppendLine("执行后原目录暂时保留；取消还原文件并移除新建分类目录。确认操作将回收以下原目录及剩余内容：");
                foreach(var dir in OriginalDirectories)lines.AppendLine(dir.Path);
                return lines.ToString();
            }
        }
    }

    internal static class FileOrganizer
    {
        internal static OrganizePlan Plan(string folder, OrganizeOptions options)
        {
            var plan = new OrganizePlan { Root = FilePaths.Folder(folder) };
            if (options.Filter && options.Types.Count == 0) throw new InvalidOperationException("启用类型筛选后，请至少勾选一种文件类型。");
            var files = FilePaths.Scan(plan.Root, true, plan.Notes);
            plan.Recycle = files.Where(f => options.Filter && !options.Types.Contains(FilePaths.Extension(f.Path))).ToList();
            files = files.Except(plan.Recycle).ToList();
            Func<FileSnapshot, string> target = f => Path.Combine(options.Classify ? Path.Combine(plan.Root, FilePaths.Extension(f.Path).Length == 0 ? "无扩展名" : FilePaths.Extension(f.Path)) : plan.Root, f.Name);
            var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            // Reserve all original names even when an earlier operation may move them away.
            foreach (var file in files) reserved.Add(file.Path);
            foreach (var file in plan.Recycle) reserved.Add(file.Path);
            if (!options.Replace)
            {
                foreach (var file in files.OrderBy(f => string.Equals(f.Path, target(f), StringComparison.OrdinalIgnoreCase) ? 0 : 1).ThenBy(f => f.Path, StringComparer.OrdinalIgnoreCase))
                {
                    string dest = target(file), basis = dest; int n = 0;
                    if (!string.Equals(file.Path, dest, StringComparison.OrdinalIgnoreCase))
                        while (reserved.Contains(dest) || File.Exists(dest) || Directory.Exists(dest)) dest = Path.Combine(Path.GetDirectoryName(basis), Path.GetFileNameWithoutExtension(basis) + "（" + (++n) + "）" + Path.GetExtension(basis));
                    reserved.Add(dest); plan.Groups.Add(new OrganizeGroup { Winner = file, Target = dest });
                }
            }
            else foreach (var group in files.GroupBy(target, StringComparer.OrdinalIgnoreCase))
            {
                var ordered = group.OrderByDescending(f => f.Modified).ThenBy(f => string.Equals(f.Path, group.Key, StringComparison.OrdinalIgnoreCase) ? 0 : 1).ThenBy(f => f.Path, StringComparer.OrdinalIgnoreCase).ToList();
                plan.Groups.Add(new OrganizeGroup { Target = group.Key, Winner = ordered[0], Backups = ordered.Skip(1).ToList() });
            }
            plan.BackupRoot = Path.Combine(plan.Root,"备份",".EI_Backup_"+Guid.NewGuid().ToString("N"));
            plan.ProtectedDirectories.Add(Path.Combine(plan.Root,"备份"));
            plan.ProtectedDirectories.AddRange(plan.Groups.Select(g=>Path.GetDirectoryName(g.Target)).Where(p=>!string.Equals(p,plan.Root,StringComparison.OrdinalIgnoreCase)).Distinct(StringComparer.OrdinalIgnoreCase));
            foreach(string dir in Directory.GetDirectories(plan.Root).OrderBy(p=>p,StringComparer.OrdinalIgnoreCase))
            {
                if(plan.ProtectedDirectories.Contains(dir,StringComparer.OrdinalIgnoreCase))continue;
                if((File.GetAttributes(dir)&FileAttributes.ReparsePoint)!=0){plan.Notes.Add("目录联接保留，不加入确认回收："+dir);continue;}
                plan.OriginalDirectories.Add(BatchDirectory.Read(dir));
            }
            return plan;
        }
        internal static FileOperationResult Execute(OrganizePlan plan, Action<string> recycle, Func<bool> cancelled, Action<string> progress, Action<string,string> move)
        {
            var result = new FileOperationResult(); move = move ?? FilePaths.Move;
            foreach(var dir in plan.OriginalDirectories)dir.Verify();
            var batch=OperationBatch.Create(plan.Root,"Organize");result.Batch=batch;result.RecoveryPath=batch.Journal;
            batch.OriginalDirectories=plan.OriginalDirectories;batch.ProtectedDirectories=plan.ProtectedDirectories;
            batch.BackupRoot=plan.BackupRoot;batch.BackupContainerCreated=!Directory.Exists(Path.Combine(plan.Root,"备份"));batch.Persist();
            foreach (string note in plan.Notes) { result.Skipped++; result.Log.AppendLine(note); }
            foreach (var item in plan.Recycle)
            {
                if (cancelled()) { result.Cancelled = true; break; }
                try { item.Verify();progress("暂存未选类型："+item.Name);string target=Path.Combine(plan.BackupRoot,item.Path.Substring(plan.Root.TrimEnd('\\').Length+1));batch.Move(batch.Track(item),target,move);result.Success++;result.Record(item.Path,"已暂存："+target); }
                catch (Exception ex) { result.Failed++; result.Record(item.Path, "暂存失败，原文件保留：" + ex.Message); }
            }
            foreach (var group in plan.Groups)
            {
                if (cancelled()) { result.Cancelled = true; break; }
                try
                {
                    progress("整理：" + group.Winner.Name); group.Winner.Verify(); foreach (var loser in group.Backups) loser.Verify();
                    FilePaths.CheckAncestors(group.Target);
                    if (Directory.Exists(group.Target)) throw new IOException("目标被文件夹占用：" + group.Target);
                    foreach (var loser in group.Backups)
                    {
                        if (cancelled()) throw new OperationCanceledException("取消同名处理；已备份文件见日志。");
                        string backup = Path.Combine(plan.BackupRoot,loser.Path.Substring(plan.Root.TrimEnd('\\').Length+1));
                        batch.Move(batch.Track(loser),backup,move); result.Success++; result.Record(loser.Path, "同名文件已备份：" + backup);
                    }
                    if (cancelled()) throw new OperationCanceledException("已取消；已备份文件见日志。");
                    if (string.Equals(group.Winner.Path, group.Target, StringComparison.OrdinalIgnoreCase)) { result.Skipped++; result.Record(group.Winner.Path, "已在目标位置，保留"); continue; }
                    batch.Move(batch.Track(group.Winner),group.Target,move);
                    result.Success++; result.Record(group.Winner.Path, "已移动：" + group.Target);
                }
                catch (OperationCanceledException ex) { result.Cancelled = true; result.Record(group.Winner.Path, "本组取消，未移动保留文件：" + ex.Message); break; }
                catch (Exception ex) { result.Failed++; result.Record(group.Winner.Path, "本组未完成，不覆盖现有目标：" + ex.Message); }
            }
            if(cancelled())result.Cancelled=true;
            batch.Log(result.Text);batch.State="Completed";batch.Persist();
            if(result.Cancelled) {var undo=batch.Undo(move);result.RecoveryIncomplete=undo.RecoveryIncomplete;result.Log.AppendLine(undo.Text);result.Success=0;}
            return result;
        }
    }
}
