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







}
