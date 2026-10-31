using System;
using System.IO;

namespace Excel_Iterate
{
    // Commit only an already closed and verified staging file. File.Replace
    // captures the displaced target as the recovery file in the same operation.
    internal static class MigrationOutputTransaction
    {
        internal static string Commit(string stage, string target, string expectedHash, string batch)
        {
            stage = Path.GetFullPath(stage); target = Path.GetFullPath(target);
            if (string.Equals(stage, target, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(Path.GetDirectoryName(stage), Path.GetDirectoryName(target), StringComparison.OrdinalIgnoreCase))
                throw new IOException("临时文件必须与目标位于同一目录，且不能是目标本身。");
            if (string.IsNullOrEmpty(batch) || batch != Path.GetFileName(batch) || batch == "." || batch == "..")
                throw new IOException("备份批次标识无效。");
            PlainPath(target); PlainPath(stage);
            if (DataMigrationHost.Digest(target) != expectedHash) throw new IOException("目标文件在预检后已变化，请重新预检。");
            // Fail before commit on ordinary sharing/read-only errors. There is
            // no delete-then-move fallback when atomic replacement is unavailable.
            using (File.Open(target, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
            string verified = DataMigrationHost.Digest(stage);
            string directory = Path.Combine(Path.GetDirectoryName(target), "迁移备份", batch);
            PlainPath(directory); Directory.CreateDirectory(directory); PlainPath(directory);
            string backup = Path.Combine(directory, Path.GetFileName(target));
            if (File.Exists(backup) || Directory.Exists(backup)) throw new IOException("备份文件已存在，未覆盖目标：" + backup);
            // Recheck after directory creation and immediately before the switch.
            if (DataMigrationHost.Digest(target) != expectedHash) throw new IOException("目标文件在提交前已变化，请重新预检。");
            File.Replace(stage, target, backup, false);
            try
            {
                if (DataMigrationHost.Digest(backup) != expectedHash) throw new IOException("提交期间目标发生并发变化。");
                if (DataMigrationHost.Digest(target) != verified) throw new IOException("替换后目标摘要与已校验内容不一致。");
            }
            catch (Exception error)
            {
                // Restore only our own installed bytes. A subsequent external
                // writer must never be overwritten by a speculative rollback.
                try
                {
                    if (DataMigrationHost.Digest(target) != verified)
                        throw new IOException("目标已被其他程序改变，未自动回滚；恢复文件：" + backup);
                    string failed = Path.Combine(directory, "未交付_" + Guid.NewGuid().ToString("N") + Path.GetExtension(target));
                    File.Replace(backup, target, failed, false);
                }
                catch (Exception rollback) { throw new IOException("提交核验失败；恢复文件位于 " + directory + "。" + rollback.Message, error); }
                throw new IOException("提交核验失败，已恢复替换前目标。", error);
            }
            return backup;
        }
        private static void PlainPath(string path)
        {
            for (string current = path; !string.IsNullOrEmpty(current); current = Path.GetDirectoryName(current))
                if ((File.Exists(current) || Directory.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("不通过重解析点替换或备份文件：" + current);
        }
    }
}
