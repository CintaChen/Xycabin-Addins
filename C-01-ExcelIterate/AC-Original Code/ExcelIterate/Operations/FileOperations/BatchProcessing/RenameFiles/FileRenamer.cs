using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Excel_Iterate
{
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
