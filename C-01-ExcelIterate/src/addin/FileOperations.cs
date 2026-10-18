using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Excel_Iterate
{
    internal sealed class FileSnapshot
    {
        internal string Path, Name, Hash; internal long Length; internal DateTime Modified;
        internal static FileSnapshot Read(string path)
        {
            var info = new FileInfo(path);
            if ((info.Attributes & FileAttributes.ReparsePoint) != 0) throw new IOException("不处理文件符号链接：" + path);
            string hash;
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(stream));
            return new FileSnapshot { Path = info.FullName, Name = info.Name, Modified = info.LastWriteTimeUtc, Length = info.Length, Hash = hash };
        }
        internal void Verify()
        {
            var now = Read(Path);
            if (now.Length != Length || now.Modified != Modified || now.Hash != Hash) throw new IOException("源文件已变化，请重新扫描：" + Path);
        }
    }
    internal sealed class FileOperationResult
    {
        internal int Success, Skipped, Failed; internal bool Cancelled, RecoveryIncomplete;
        internal string RecoveryPath; internal OperationBatch Batch;
        internal readonly StringBuilder Log = new StringBuilder();
        internal readonly Dictionary<string, string> Status = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        internal void Record(string source, string status) { Status[source] = status; Log.AppendLine(source + "：" + status); }
        internal string Summary { get { return "成功 " + Success + " 项，跳过 " + Skipped + " 项，失败 " + Failed + " 项。" + (Cancelled ? "已取消，还原结果见日志。" : "") + (RecoveryIncomplete ? "还原未完成，请依据恢复记录处理。" : ""); } }
        internal string Text { get { return Summary + "\r\n\r\n" + Log + (RecoveryPath == null ? "" : "\r\n恢复记录：" + RecoveryPath); } }
    }
    internal static class FilePaths
    {
        internal static string Folder(string input)
        {
            string text = (input ?? "").Trim().Trim('"');
            if (!System.IO.Path.IsPathRooted(text)) throw new InvalidOperationException("请填写完整文件夹路径。");
            string root = System.IO.Path.GetFullPath(text).TrimEnd('\\', '/');
            if (root.Length == 2 && root[1] == ':') root += "\\";
            if (!Directory.Exists(root)) throw new DirectoryNotFoundException("文件夹不存在：" + root);
            CheckAncestors(root); return root;
        }
        internal static void CheckAncestors(string path)
        {
            string current = Directory.Exists(path) ? path : System.IO.Path.GetDirectoryName(path);
            while (!string.IsNullOrEmpty(current))
            {
                if (Directory.Exists(current) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0) throw new IOException("不跟随目录联接或符号链接：" + current);
                current = System.IO.Path.GetDirectoryName(current);
            }
        }
        internal static string Extension(string path) { return System.IO.Path.GetExtension(path).TrimStart('.').ToLowerInvariant(); }
        internal static List<FileSnapshot> Scan(string root, bool recursive, List<string> notes)
        {
            root = Folder(root); var result = new List<FileSnapshot>(); var pending = new Stack<string>(); pending.Push(root);
            while (pending.Count > 0)
            {
                string dir = pending.Pop(); CheckAncestors(dir);
                foreach (string file in Directory.GetFiles(dir).OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
                {
                    if (OperationBatch.ManagedFile(file)) continue;
                    if ((File.GetAttributes(file) & FileAttributes.ReparsePoint) != 0) { notes.Add("跳过文件符号链接：" + file); continue; }
                    result.Add(FileSnapshot.Read(file));
                }
                if (recursive) foreach (string child in Directory.GetDirectories(dir).OrderByDescending(p => p, StringComparer.OrdinalIgnoreCase))
                {
                    if(string.Equals(dir,root,StringComparison.OrdinalIgnoreCase)&&string.Equals(Path.GetFileName(child),"备份",StringComparison.OrdinalIgnoreCase))continue;
                    if (Path.GetFileName(child).StartsWith(".EI_Backup_", StringComparison.OrdinalIgnoreCase) && Path.GetFileName(dir)=="备份") continue;
                    if ((File.GetAttributes(child) & FileAttributes.ReparsePoint) != 0) { notes.Add("跳过目录联接：" + child); continue; }
                    pending.Push(child);
                }
            }
            return result;
        }
        internal static void Move(string source, string target)
        { CheckAncestors(source); CheckAncestors(target); File.Move(source, target); }
    }
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
    internal sealed class RenameItem
    {
        internal FileSnapshot Source; internal string Input, Target, Current, Temp;
    }
    internal sealed class RenamePlan
    {
        internal string Root, DocumentPath, Anchor; internal List<RenameItem> Items = new List<RenameItem>();
    }
    internal static class FileRenamer
    {
        internal static string NewName(string input, string original)
        {
            string name = (input ?? "").Trim();
            if (name.Length == 0) return null;
            if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || name == "." || name == ".." || name.EndsWith(".") || name.EndsWith(" ")) throw new InvalidOperationException("名称含非法字符或结尾：" + name);
            string stem = name.Split('.')[0].TrimEnd(' ');
            if (System.Text.RegularExpressions.Regex.IsMatch(stem, @"^(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) throw new InvalidOperationException("Windows 保留名称：" + name);
            if (Path.GetExtension(name).Length == 0) name += Path.GetExtension(original);
            if (name.Length > 255) throw new InvalidOperationException("文件名过长：" + name);
            return name;
        }
        internal static FileOperationResult Preflight(RenamePlan plan)
        {
            var result = new FileOperationResult(); var targets = new Dictionary<string, RenameItem>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in plan.Items)
            {
                item.Target = null; item.Current = item.Source.Path;
                try
                {
                    FilePaths.CheckAncestors(item.Source.Path); item.Source.Verify();
                    if (!string.Equals(Path.GetDirectoryName(item.Source.Path), plan.Root, StringComparison.OrdinalIgnoreCase)) throw new IOException("源文件越出所选目录。");
                    string name = NewName(item.Input, item.Source.Name);
                    if (name == null || name == item.Source.Name) { result.Skipped++; result.Record(item.Source.Path, name == null ? "未填新名称，跳过" : "名称相同，无需修改"); continue; }
                    item.Target = Path.Combine(plan.Root, name);
                    if (!string.Equals(Path.GetDirectoryName(item.Target), plan.Root, StringComparison.OrdinalIgnoreCase)) throw new IOException("目标越出所选目录。");
                    RenameItem previous;
                    if (targets.TryGetValue(item.Target, out previous)) { result.Failed++; result.Record(previous.Source.Path, "目标名称重复：" + name); throw new IOException("目标名称重复：" + name); }
                    targets.Add(item.Target, item);
                    using (File.Open(item.Source.Path, FileMode.Open, FileAccess.Read, FileShare.None)) { }
                }
                catch (Exception ex) { result.Failed++; result.Record(item.Source.Path, "预检失败：" + ex.Message); }
            }
            var participating = new HashSet<string>(plan.Items.Where(i => i.Target != null).Select(i => i.Source.Path), StringComparer.OrdinalIgnoreCase);
            foreach (var item in plan.Items.Where(i => i.Target != null))
                if (Directory.Exists(item.Target) || (File.Exists(item.Target) && !participating.Contains(item.Target))) { result.Failed++; result.Record(item.Source.Path, "目标被未参与改名的文件或目录占用：" + item.Target); }
            if (result.Failed > 0) result.Log.AppendLine("整批暂停；未改动任何文件。修正后重新执行。");
            return result;
        }
        internal static FileOperationResult Execute(RenamePlan plan, Func<bool> cancelled, Action<string> progress, Action<string,string> move)
        {
            var result = Preflight(plan); if (result.Failed > 0) return result;
            move = move ?? FilePaths.Move; var active = plan.Items.Where(i => i.Target != null).ToList();
            if (active.Count == 0) return result;
            var batch=OperationBatch.Create(plan.Root,"Rename");result.Batch=batch;result.RecoveryPath=batch.Journal;
            batch.DocumentPath=plan.DocumentPath;batch.Anchor=plan.Anchor;batch.ReportFiles=plan.Items.Select(i=>BatchFile.From(i.Source)).ToList();batch.Persist();
            var entries=active.ToDictionary(i=>i.Source.Path,i=>batch.Track(i.Source),StringComparer.OrdinalIgnoreCase);
            Action persist = delegate { batch.Persist(); };
            try
            {
                foreach (var item in active) item.Temp = Path.Combine(plan.Root, ".EI_Rename_" + Guid.NewGuid().ToString("N") + ".tmp");
                persist();
                foreach (var item in active)
                {
                    if (cancelled()) throw new OperationCanceledException();
                    item.Source.Verify(); progress("暂存：" + item.Source.Name);
                    batch.Move(entries[item.Source.Path],item.Temp,move); item.Current = item.Temp;
                }
                foreach (var item in active)
                {
                    if (cancelled()) throw new OperationCanceledException();
                    progress("改名：" + Path.GetFileName(item.Target));batch.Move(entries[item.Source.Path],item.Target,move); item.Current = item.Target;
                    result.Success++; result.Record(item.Source.Path, "已改名：" + Path.GetFileName(item.Target));
                }
                batch.State="Completed";persist();result.Log.AppendLine("整批提交完成；取消操作可还原，确认操作清理记录。");
            }
            catch (Exception ex)
            {
                result.Cancelled = ex is OperationCanceledException; result.Success = 0; result.Failed++;
                result.Log.AppendLine("提交中止：" + ex.Message + "；尝试整批还原。");
                var undo=batch.Undo(move);result.RecoveryIncomplete=undo.RecoveryIncomplete;result.Log.AppendLine(undo.Text);
                foreach(var entry in batch.Files)result.Record(entry.Source,entry.Current==entry.Source?"本批未提交，已还原并核对":"还原未完成；当前路径："+entry.Current);
            }
            batch.Log(result.Text);return result;
        }
    }
}
