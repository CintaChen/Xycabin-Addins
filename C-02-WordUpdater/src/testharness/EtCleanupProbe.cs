using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

class EtCleanupProbe
{
    static int checks;
    static void Check(bool value, string name)
    { if (!value) throw new Exception(name); checks++; }
    static int Main(string[] args)
    {
        try
        {
            if (args.Length > 1) throw new ArgumentException("Usage: EtCleanupProbe.exe [WordUpdater.dll]");
            string assemblyPath = args.Length == 1 ? Path.GetFullPath(args[0]) : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WordUpdater.dll");
            Assembly assembly = Assembly.LoadFrom(assemblyPath);
            MethodInfo review = assembly.GetType("WordUpdater.EtCleanup", true).GetMethod("ReviewCandidates", BindingFlags.NonPublic | BindingFlags.Static);
            var inspected = new List<int>();
            var logs = new List<string>();
            Func<HashSet<int>, int[], Func<int, bool?>, int> run = (before, now, visible) =>
                (int)review.Invoke(null, new object[] { before, now, visible, new Action<string>(logs.Add) });
            Func<int, bool?> states = pid => { inspected.Add(pid); return pid == 20 ? (bool?)true : pid == 30 ? null : (bool?)false; };
            Check(run(new HashSet<int> { 10 }, new[] { 10, 20, 30, 40, 40 }, states) == 0, "No process termination counted");
            Check(!inspected.Contains(10), "Existing instance preserved");
            Check(inspected.Count == 3, "Duplicate concurrent PID reviewed once");
            Check(logs.Count == 2, "Visible instance has no hidden-candidate claim");
            Check(logs[0].Contains("无法确认") && logs[0].Contains("pid=30"), "Uncertain enumeration diagnosed");
            Check(logs[1].Contains("隐藏ET") && logs[1].Contains("pid=40"), "Concurrent hidden process retained");
            Check(logs.TrueForAll(text => text.Contains("未发送关闭或终止请求")), "PID difference never establishes ownership");
            inspected.Clear(); logs.Clear();
            Check(run(null, new[] { 40 }, states) == 0 && inspected.Count == 0, "Missing baseline fails closed");
            Check(logs.Count == 1 && logs[0].Contains("快照不可用"), "Missing baseline diagnosed");
            logs.Clear();
            Check(run(new HashSet<int>(), new[] { 50 }, pid => { throw new InvalidOperationException("injected"); }) == 0, "Window failure preserves instance");
            Check(logs.Count == 1 && logs[0].Contains("injected") && logs[0].Contains("保留"), "Failure evidence retained");
            Console.WriteLine("EtCleanupProbe: " + checks + " PASS, 0 FAIL; fake candidates only, no WPS attached");
            return 0;
        }
        catch (Exception error) { Console.WriteLine(error); return 1; }
    }
}
