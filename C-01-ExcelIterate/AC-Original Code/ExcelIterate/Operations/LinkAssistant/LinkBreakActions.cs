using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Excel_Iterate
{
    internal sealed partial class LinkAssistantSession
    {
        private sealed class CheckpointOutcome { internal int Success; internal bool Cancelled; internal Exception Error; }
        private int PendingChanges()
        {
            var owners = new HashSet<string>(Batch.Checkpoints.Where(c => c.NeedsRestore).Select(c => c.Owner), StringComparer.OrdinalIgnoreCase);
            return Batch.Changes.Count(c => c.Applied || c.Intent || owners.Contains(c.Owner));
        }
        private LinkWorkbookCheckpoint Checkpoint(string owner) { return Batch.Checkpoints.FirstOrDefault(c => string.Equals(c.Owner, owner, StringComparison.OrdinalIgnoreCase)); }
        private void GuardCheckpoint(LinkWorkbookCheckpoint checkpoint, object book)
        {
            if (book != null && !ComFormulaHost.Boolean(((dynamic)book).Saved, "Saved")) throw new IOException("断链备份后工作簿有新编辑，已停止；请保留备份并核对，不会自动覆盖新编辑。");
            if (XlsxLinkPackage.Hash(checkpoint.Owner) != checkpoint.LastHash) throw new IOException("断链备份后文件已变化，已停止并保留备份。");
        }
        private LinkWorkbookCheckpoint CreateCheckpoint(string owner)
        {
            var checkpoint = Checkpoint(owner); if (checkpoint != null) { Batch.ValidateBackup(checkpoint); return checkpoint; }
            if ((File.GetAttributes(owner) & (FileAttributes.ReadOnly | FileAttributes.ReparsePoint)) != 0) throw new IOException("断链目标为只读文件或重解析点，未修改：" + owner);
            string hash = XlsxLinkPackage.Hash(owner);
            checkpoint = new LinkWorkbookCheckpoint { Owner = owner, BeforeHash = hash, LastHash = hash,
                Backup = Path.Combine(Path.GetDirectoryName(owner), ".EI_LinkBackup_" + Batch.Id + "_" + Guid.NewGuid().ToString("N") + ".xlsx") };
            File.Copy(owner, checkpoint.Backup, false);
            Batch.Checkpoints.Add(checkpoint); Batch.Persist();
            Batch.ValidateBackup(checkpoint); GuardCheckpoint(checkpoint, null);
            Operations.Record("断链备份", owner + "\r\n" + checkpoint.Backup);
            return checkpoint;
        }
        private static HashSet<string> SavedSources(string path)
        { return new HashSet<string>(XlsxLinkPackage.Scan(path).Where(l => l.SourcePath != null).Select(l => l.SourcePath), StringComparer.OrdinalIgnoreCase); }
        private void RestoreCheckpoint(LinkWorkbookCheckpoint checkpoint)
        {
            if (checkpoint == null) return;
            if (!checkpoint.NeedsRestore && !Batch.Changes.Any(c => c.Owner == checkpoint.Owner && (c.Applied || c.Intent))) { checkpoint.Restored = true; Batch.Persist(); return; }
            Batch.ValidateBackup(checkpoint);
            using (var bag = new ComBag())
                if (bag.Own(FindBook(App, checkpoint.Owner)) != null) throw new IOException("请先关闭待恢复工作簿，再点击「取消操作」：" + checkpoint.Owner + "。备份已保留；不会自动关闭工作簿。");
            string current = XlsxLinkPackage.Hash(checkpoint.Owner);
            if (checkpoint.CommitIntent && current == checkpoint.PreparedHash)
            { checkpoint.LastHash = current; checkpoint.Applied = true; foreach (var c in Batch.Changes.Where(c => c.Owner == checkpoint.Owner)) c.EverApplied = true; }
            if (current != checkpoint.LastHash) throw new IOException("断链后工作簿有新增改动或提交状态不明确，未覆盖；请保留备份核对：" + checkpoint.Backup);
            if (current != checkpoint.BeforeHash)
            {
                string stage = Path.Combine(Path.GetDirectoryName(checkpoint.Owner), ".EI_LinkStage_" + Guid.NewGuid().ToString("N") + ".xlsx");
                try
                {
                    File.Copy(checkpoint.Backup, stage, false);
                    if (XlsxLinkPackage.Hash(stage) != checkpoint.BeforeHash || XlsxLinkPackage.Hash(checkpoint.Owner) != current) throw new IOException("恢复提交前摘要变化，未覆盖原文件。");
                    File.Replace(stage, checkpoint.Owner, null);
                    if (XlsxLinkPackage.Hash(checkpoint.Owner) != checkpoint.BeforeHash) throw new IOException("恢复后摘要核对失败，备份已保留。");
                }
                finally { if (File.Exists(stage)) File.Delete(stage); }
            }
            foreach (var c in Batch.Changes.Where(c => c.Owner == checkpoint.Owner)) { c.Applied = c.Intent = false; }
            checkpoint.Applied = checkpoint.Uncertain = checkpoint.CommitIntent = false; checkpoint.Restored = true; checkpoint.LastHash = checkpoint.BeforeHash; Batch.Persist();
        }
        private IEnumerable<string> CheckpointSteps(string owner, bool undo, Func<bool> cancelled, CheckpointOutcome outcome)
        {
            if (undo)
            {
                try
                {
                    var checkpoint = Checkpoint(owner);
                    if (checkpoint == null && Batch.Changes.Any(c => c.Owner == owner && (c.Applied || c.Intent))) throw new IOException("断链恢复记录缺少整簿备份，请保留证据人工核对。");
                    var restored = Batch.Changes.Where(c => c.Owner == owner && (c.Applied || c.Intent || c.EverApplied)).ToList();
                    RestoreCheckpoint(checkpoint);
                    foreach (var c in restored) Pair(c, true, true, null);
                    outcome.Success = restored.Count;
                }
                catch (Exception error) { outcome.Error = error; }
                yield return outcome.Error == null ? "已核对整簿恢复：" + Path.GetFileName(owner) : "断链备份已保留：" + Path.GetFileName(owner);
                yield break;
            }
            LinkWorkbookCheckpoint snapshot = null; List<LinkChange> ordered = null;
            try
            {
                using (var bag = new ComBag()) { var open = bag.Own(FindBook(App, owner)); if (open != null) RequireSaved(open); }
                var actual = SavedSources(owner);
                var changes = Batch.Changes.Where(c => c.Owner == owner).ToList();
                foreach (var c in changes.Where(c => c.Action == LinkAction.Replace).ToArray()) if (MissingTarget(c.New)) { SkipMissing(c); changes.Remove(c); }
                if (changes.Any(c => !actual.Contains(c.Old))) throw new IOException("链接来源已变化，未断开：" + owner);
                ordered = LinkBatch.Order(changes.Where(c => c.Action == LinkAction.Replace), actual, false);
                ordered.AddRange(changes.Where(c => c.Action == LinkAction.Break));
                var names = LinkNameAnalysis.Read(owner);
                foreach (var change in changes.Where(c => c.Action == LinkAction.Break)) names.Plan(change.Old);
                snapshot = CreateCheckpoint(owner);
            }
            catch (Exception error) { outcome.Error = error; }
            if(outcome.Error!=null)yield break;
            yield return "备份及断链预检完成：" + Path.GetFileName(owner);
            if (cancelled()) { outcome.Cancelled = true; yield break; }
            using (var bag = new ComBag())
            {
                object original = bag.Own(FindBook(App, owner)), work = original, active = null;
                string stage = null,stageHash=null; bool ownCopy = false,preserveCopy=false;
                try
                {
                    try
                    {
                        GuardCheckpoint(snapshot, original);
                        if (original == null)
                        {
                            stage = Path.Combine(Path.GetDirectoryName(owner), ".EI_LinkStage_" + Guid.NewGuid().ToString("N") + ".xlsx");
                            File.Copy(owner, stage, false);
                            active = bag.Own(((dynamic)App).ActiveWorkbook);
                            dynamic books = bag.Own(((dynamic)App).Workbooks);
                            using (var quiet = new HostQuiet(App)) work = bag.Own(books.Open(stage, 0, false));
                            if (work == null || !string.Equals(Path.GetFullPath(Convert.ToString(((dynamic)work).FullName)), stage, StringComparison.OrdinalIgnoreCase)) throw new IOException("宿主没有返回本次暂存副本，已停止；不会关闭其他工作簿。");
                            ownCopy = true;
                            stageHash=XlsxLinkPackage.Hash(stage);
                        }
                    }
                    catch (Exception error) { outcome.Error = error; }
                    if (outcome.Error != null) yield break;
                    foreach (var change in ordered)
                    {
                        if (cancelled()) { outcome.Cancelled = true; break; }
                        try
                        {
                            if (original != null) GuardCheckpoint(snapshot, work); else GuardCheckpoint(snapshot, null);
                            if(original==null&&(!ComFormulaHost.Boolean(((dynamic)work).Saved,"Saved")||XlsxLinkPackage.Hash(stage)!=stageHash))
                            {preserveCopy=true;throw new IOException("暂存副本有新增编辑，已保留副本且未提交原文件："+stage);}
                            string path = original == null ? stage : owner;
                            var sources = ActiveHostSources(work, path, s => Operations.Trace("断链宿主来源", s));
                            var expected = new HashSet<string>(sources.Select(s => s.Path), StringComparer.OrdinalIgnoreCase);
                            var source = sources.SingleOrDefault(s => string.Equals(s.Path, change.Old, StringComparison.OrdinalIgnoreCase));
                            if (source == null) throw new IOException("不能唯一定位待处理来源：" + change.Old);
                            if (change.Action == LinkAction.Replace && (expected.Contains(change.New) || MissingTarget(change.New))) throw new IOException("执行期间替换目标变化，已停止该工作簿。");
                            LinkNameAnalysis nameAnalysis = null; List<LinkDefinedName> namePlan = null;
                            List<object> nativeNames = null;
                            using (var nameReferences = new ComBag())
                            {
                                if (change.Action == LinkAction.Break)
                                {
                                    nameAnalysis = LinkNameAnalysis.Read(path); namePlan = nameAnalysis.Plan(change.Old);
                                    nativeNames = nameAnalysis.NativeNames(work, namePlan, nameReferences);
                                }
                                GuardCheckpoint(snapshot, work);
                                if (original == null && XlsxLinkPackage.Hash(stage) != stageHash)
                                { preserveCopy = true; throw new IOException("名称读取期间暂存副本发生新编辑，已保留副本且未提交原文件。"); }
                                if (original != null) { change.Intent = true; snapshot.Uncertain = true; Batch.Persist(); }
                                using (var quiet = new HostQuiet(App))
                                {
                                    if (change.Action == LinkAction.Break)
                                    {
                                        foreach (object name in nativeNames) ((dynamic)name).Delete();
                                        if (namePlan.Count > 0) Operations.Trace("清理来源名称", change.Old + "；" + namePlan.Count + " 个未使用名称");
                                        // Name deletion can remove the native source immediately;
                                        // match fresh raw names without requiring the old package to agree.
                                        var current = ParseHostSources((object)((dynamic)work).LinkSources(1), path, XlsxLinkPackage.Scan(path))
                                            .Where(s => string.Equals(s.Path, change.Old, StringComparison.OrdinalIgnoreCase)).GroupBy(s => s.Name, StringComparer.Ordinal).Select(g => g.First()).ToList();
                                        if (current.Count > 1) throw new IOException("清理名称后宿主来源不唯一，备份已保留。");
                                        if (current.Count == 1) { Operations.Trace("调用BreakLink", "Name=" + current[0].Name + "；Type=1"); ((dynamic)work).BreakLink(current[0].Name, 1); }
                                    }
                                    else ((dynamic)work).ChangeLink(source.Name, change.New, 1);
                                    WorkbookSaving.Save(work, true);
                                }
                                if (original != null) { snapshot.LastHash = XlsxLinkPackage.Hash(owner); snapshot.Applied = true; snapshot.Uncertain = false; change.EverApplied |= snapshot.LastHash != snapshot.BeforeHash; Batch.Persist(); }
                                else stageHash=XlsxLinkPackage.Hash(stage);
                                expected.Remove(change.Old); if (change.Action == LinkAction.Replace) expected.Add(change.New);
                                if (!expected.SetEquals(SavedSources(path))) throw new IOException("保存后断链或剩余来源核对失败，备份已保留。");
                                if (nameAnalysis != null) nameAnalysis.Verify(LinkNameAnalysis.Read(path), change.Old, namePlan);
                                if (original != null) { change.Applied = change.EverApplied = true; change.Intent = false; Batch.Persist(); outcome.Success++; Pair(change, false, true, null); }
                            }
                        }
                        catch (Exception error) { outcome.Error = error; break; }
                        yield return (original == null ? "已处理暂存副本：" : "已保存链接操作：") + Path.GetFileName(owner);
                    }
                    if (original == null && outcome.Error == null && !outcome.Cancelled)
                    {
                        try
                        {
                            if (cancelled()) { outcome.Cancelled = true; }
                            else
                            {
                                if(!ComFormulaHost.Boolean(((dynamic)work).Saved,"Saved")||XlsxLinkPackage.Hash(stage)!=stageHash)
                                {preserveCopy=true;throw new IOException("暂存副本提交前出现新增编辑，已保留副本且未提交原文件："+stage);}
                                ((dynamic)work).Close(false); ownCopy = false;
                                GuardCheckpoint(snapshot, null);
                                snapshot.PreparedHash = XlsxLinkPackage.Hash(stage); snapshot.CommitIntent = true;
                                foreach (var c in ordered) c.Intent = true; Batch.Persist();
                                File.Replace(stage, owner, null); stage = null;
                                if (XlsxLinkPackage.Hash(owner) != snapshot.PreparedHash) throw new IOException("暂存提交后摘要核对失败。");
                                snapshot.LastHash = snapshot.PreparedHash; snapshot.CommitIntent = false; snapshot.Applied = true;
                                foreach (var c in ordered) { c.Applied = c.EverApplied = true; c.Intent = false; }
                                Batch.Persist(); outcome.Success = ordered.Count;
                                foreach (var c in ordered) Pair(c, false, true, null);
                            }
                        }
                        catch (Exception error) { outcome.Error = error; }
                    }
                }
                finally
                {
                    if (ownCopy && !preserveCopy)
                    {
                        try
                        {
                            if (!ComFormulaHost.Boolean(((dynamic)work).Saved, "Saved") || stage == null || XlsxLinkPackage.Hash(stage) != stageHash)
                            { preserveCopy = true; Operations.Record("暂存副本已保留", "副本存在未保存编辑或内容变化，未关闭、未提交原文件：" + stage); }
                        }
                        catch (Exception error) { preserveCopy = true; if (outcome.Error == null) outcome.Error = error; Operations.Trace("暂存副本核对失败", stage + "；" + error); }
                    }
                    if (ownCopy&&!preserveCopy) try { ((dynamic)work).Close(false); } catch (Exception error) { if (outcome.Error == null) outcome.Error = error;preserveCopy=true; }
                    if (active != null) try { ((dynamic)active).Activate(); } catch { }
                    if (stage != null && File.Exists(stage)&&!preserveCopy) try { File.Delete(stage); } catch (Exception error) { Operations.Trace("暂存副本保留", stage + "；" + error.GetType().Name); }
                }
            }
        }
    }
}
