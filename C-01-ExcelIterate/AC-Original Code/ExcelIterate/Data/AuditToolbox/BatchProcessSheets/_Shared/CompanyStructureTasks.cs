using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    internal sealed class CompanyStructureTask
    {
        internal int Row;
        internal string Sequence, ShortName, FullName, FileName, Level;
        internal bool Leaf;
    }
    internal sealed class CompanyStructureColumns
    {
        internal int FirstRow = 3, Sequence = 1, ShortName = 2, FullName = 4, FileName = 6, Level = 3;
    }
    internal static class CompanyStructureTasks
    {
        internal static string Key(string sequence)
        {
            string text = (sequence ?? "").Trim();
            if (!Regex.IsMatch(text, @"^[0-9]+(?:-[0-9]+)*-?$")) throw new InvalidOperationException("序号格式不正确，请使用03-、03-01-这样的分段数字。");
            return string.Join("-", text.TrimEnd('-').Split('-').Select(s => s.TrimStart('0').Length == 0 ? "0" : s.TrimStart('0')).ToArray());
        }
        internal static void Classify(IList<CompanyStructureTask> rows, bool useLevel)
        {
            var keys = new Dictionary<string, CompanyStructureTask>(StringComparer.Ordinal);
            foreach (var row in rows)
            {
                string key;
                try { key = Key(row.Sequence); }
                catch (Exception error) { throw new InvalidOperationException("架构表第" + row.Row + "行：" + error.Message); }
                if (keys.ContainsKey(key)) throw new InvalidOperationException("架构表序号重复：第" + keys[key].Row + "行与第" + row.Row + "行，" + row.Sequence);
                keys.Add(key, row);
            }
            foreach (var entry in keys)
                entry.Value.Leaf = !keys.Keys.Any(k => k.StartsWith(entry.Key + "-", StringComparison.Ordinal));
            if (useLevel)
            {
                var roots = rows.Where(r => Regex.IsMatch((r.Level ?? "").Trim(), @"^0\s*级?$" )).ToList();
                if (roots.Count > 1) throw new InvalidOperationException("架构表存在多个0级总集团，无法确定归属，请核对级次列。");
                if (roots.Count == 1 && rows.Count > 1) roots[0].Leaf = false;
            }
        }
        internal static List<CompanyWorkbookRecord> Records(IEnumerable<CompanyStructureTask> tasks)
        {
            return tasks.Select(t => new CompanyWorkbookRecord { Row = t.Row, Name = t.FileName, Code = t.ShortName ?? "", FullName = t.FullName }).ToList();
        }
        internal static List<string> ResolveFiles(IEnumerable<CompanyStructureTask> tasks,string folder)
        {
            var result=new List<string>();var names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(var task in tasks)
            {
                string file;
                try{file=CompanyWorkbookRules.FileName(task.FileName);}
                catch(Exception error){throw new InvalidOperationException("架构表第"+task.Row+"行："+error.Message,error);}
                if(!names.Add(file))throw new InvalidOperationException("架构表文件名重复："+file);
                string extension=Path.GetExtension(file);if(!string.Equals(extension,".xlsx",StringComparison.OrdinalIgnoreCase)&&!string.Equals(extension,".xls",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("架构表第"+task.Row+"行文件格式暂不支持："+file);
                result.Add(Path.Combine(Path.GetFullPath(folder),file));
            }
            if(result.Count==0)throw new InvalidOperationException("请至少勾选一个架构任务。");return result;
        }
    }
}
