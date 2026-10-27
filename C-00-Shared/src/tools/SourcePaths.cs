using System;
using System.IO;
using System.Linq;
using System.Text;

namespace WpsShared
{
    // Source assertions must read the implementations actually compiled, including
    // linked common files. Non-source fixture/output paths retain normal IO behavior.
    internal static class SourcePaths
    {
        internal static string Resolve(string path)
        {
            string full = Path.GetFullPath(path);
            foreach (string product in new[] { "C-01-ExcelIterate", "C-02-WordUpdater" })
            {
                string marker = Path.DirectorySeparatorChar + product + Path.DirectorySeparatorChar;
                int offset = full.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                if (offset < 0) continue;
                string common = Path.Combine(full.Substring(0, offset), "C-00-Shared", full.Substring(offset + marker.Length));
                if (File.Exists(common) && !File.Exists(full)) return common;
            }
            return full;
        }
        internal static bool Exists(string path) { return File.Exists(Resolve(path)); }
        internal static string ReadAllText(string path) { return File.ReadAllText(Resolve(path)); }
        internal static string ReadAllText(string path, Encoding encoding) { return File.ReadAllText(Resolve(path), encoding); }
        internal static string[] GetFiles(string path, string pattern, SearchOption option)
        {
            string full = Path.GetFullPath(path);
            var local = Directory.Exists(full) ? Directory.GetFiles(full, pattern, option) : new string[0];
            foreach (string product in new[] { "C-01-ExcelIterate", "C-02-WordUpdater" })
            {
                string marker = Path.DirectorySeparatorChar + product + Path.DirectorySeparatorChar;
                int offset = full.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                if (offset < 0 || !full.Substring(offset + marker.Length).StartsWith("src", StringComparison.OrdinalIgnoreCase)) continue;
                string common = Path.Combine(full.Substring(0, offset), "C-00-Shared", full.Substring(offset + marker.Length));
                if (Directory.Exists(common)) local = local.Concat(Directory.GetFiles(common, pattern, option)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            }
            return local;
        }
        internal static string[] GetFiles(string path, string pattern) { return GetFiles(path, pattern, SearchOption.TopDirectoryOnly); }
        internal static string[] GetFiles(string path) { return GetFiles(path, "*"); }
    }
}
