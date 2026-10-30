using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel_Iterate;

// Explicit integration test. All documents are created in ProbeSession's own
// directory. Never install the add-in, call Quit, or write an existing document.
internal static class DataMigrationLiveProbe
{
    private static void Check(bool passed, string message)
    { if (!passed) throw new InvalidOperationException(message); Console.WriteLine("PASS: " + message); }
    private static void Drain(IEnumerable<MigrationProgress> steps)
        { foreach (var step in steps) Console.WriteLine("STEP: " + step.Message); }
    private static void DrainPumped(IEnumerable<MigrationProgress> sequence)
    {
        Exception failure = null; bool done = false, ticking = false;
        using (var context = new ApplicationContext()) using (var timer = new Timer { Interval = 15 }) using (var steps = sequence.GetEnumerator())
        {
            timer.Tick += delegate { if (ticking) return; ticking = true; timer.Stop(); try { if (!steps.MoveNext()) { done = true; context.ExitThread(); } else Console.WriteLine("PUMP: " + steps.Current.Message); } catch (Exception error) { failure = error; done = true; context.ExitThread(); } finally { ticking = false; if (!done) timer.Start(); } };
            timer.Start(); Application.Run(context);
        }
        if (failure != null) throw failure;
    }
    private static void Reproduce(ProbeSession session, object application, string[] args)
    {
        dynamic app = application;
        string source = session.Output("repro-source.xlsx"); File.Copy(args[1], source);
        var targets = new List<MigrationTarget>();
        for (int i = 2; i <= 3; i++)
        {
            string path = session.Output(Path.GetFileName(args[i])); File.Copy(args[i], path);
            targets.Add(new MigrationTarget { Path = path, Sheet = "Sheet1", Start = args[0] == "--repro" ? "A1" : "B2" });
            using (var refs = session.References())
            {
                dynamic books = refs.Take(app.Workbooks); object owned = books.Open(path, 0, false);
                session.Own(owned, delegate { ((dynamic)owned).Close(false); });
            }
        }
        session.Override(() => (object)app.EnableEvents, value => app.EnableEvents = (dynamic)value, true, "repro events");
        DataMigrationHost.DiagnosticRootOverride = Path.Combine(session.DirectoryPath, "diagnostics");
        using (var host = new DataMigrationHost(application))
        {
            if (args[0] != "--repro")
            {
                bool recalculated = false;
                var request = new MigrationRequest { SourcePath = source, SourceSheet = "迁移源", SourceRange = "B2:I22", Targets = targets };
                RunFormBatch(host, request, delegate(MigrationProgress step) {
                    if (args[0] != "--diagnose" || recalculated || !step.Message.StartsWith("写入 " + Path.GetFileName(targets[1].Path) + " ", StringComparison.Ordinal)) return;
                    recalculated = true;
                    var handle = (MergeWorkbookHandle)typeof(DataMigrationHost).GetField("source", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(host);
                    Console.WriteLine("DIAGNOSTIC: calculate source sheet without editing any values or formulas");
                    ((dynamic)handle.Sheet).Calculate();
                });
                Console.WriteLine("EVENT_TRACE:\n"+host.EventTrace);
            }
            else { DrainPumped(host.Preview(new MigrationRequest { SourcePath = source, SourceSheet = "迁移源", SourceRange = "A1", AllowEmptySource = true, Targets = targets })); DrainPumped(host.Execute()); }
            foreach (var target in targets) { Console.WriteLine("TARGET: " + target.Path + " " + target.Status + " " + target.Error + "\n" + target.FailureDetails); Check(target.Status == "成功", "two open target workbooks processed independently"); if (target.Output.Length > 0) session.TrackFile(target.Output); }
            if(args[0]=="--repair")foreach(var target in targets){string output=Path.Combine(Path.GetFullPath(args[4]),Path.GetFileNameWithoutExtension(target.Path)+"_迁移修正结果.xlsx");session.Export(target.Output,output);Console.WriteLine("REPAIRED_OUTPUT: "+output);}
        }
        foreach (string file in Directory.GetFiles(session.DirectoryPath, "*", SearchOption.AllDirectories)) session.TrackFile(file);
    }
    private static IEnumerable<MigrationProgress> Observed(IEnumerable<MigrationProgress> steps, Action<MigrationProgress> progress)
    { foreach (var step in steps) { if (progress != null) progress(step); yield return step; } }
    private static void RunFormBatch(DataMigrationHost host, MigrationRequest request, Action<MigrationProgress> progress = null)
    {
        Exception failure = null;
        using (var context = new ApplicationContext()) using (var form = new DataMigrationForm())
        {
            form.SetSource(request.SourcePath, request.SourceSheet, request.SourceRange, new[] { request.SourceSheet });
            foreach (var target in request.Targets) { form.AddTarget(target.Path, new[] { target.Sheet }, target.Sheet, target.Start); form.Targets.Rows[form.Targets.Rows.Count - 1].Tag = target; }
            Action<Exception> failed = delegate(Exception error) { failure = error; context.ExitThread(); };
            form.Run(host.Preview(request), delegate {
                form.RefreshResults();
                form.Run(Observed(host.Execute(),progress), delegate { form.RefreshResults(); Console.WriteLine("FORM_SUMMARY: " + DataMigrationForm.Summary(request)); context.ExitThread(); }, failed);
            }, failed);
            Application.Run(context);
        }
        if (failure != null) throw failure;
    }
    private static void Make(ProbeSession session, object app, string path, int format, string sheetName, Action<object, object> populate)
    {
        using (var refs = session.References())
        {
            dynamic books = refs.Take(((dynamic)app).Workbooks); object created = null; bool closed = false;
            try
            {
                created = session.Own((object)books.Add(), delegate { if (!closed) { ((dynamic)created).Close(false); closed = true; } });
                dynamic book = created; dynamic sheets = refs.Take(book.Worksheets); object sheet = refs.Take(sheets.Item(1));
                ((dynamic)sheet).Name = sheetName; populate(created, sheet); book.SaveAs(path, format);
                book.Close(false); closed = true;
            }
            finally { if (!closed && created != null) { ((dynamic)created).Close(false); closed = true; } }
        }
    }
    private static void Cell(ProbeSession session, object sheet, string address, object value, bool formula)
    {
        using (var refs = session.References())
        {
            dynamic cell = refs.Take(((dynamic)sheet).Range[address]);
            if (formula) cell.Formula = value; else { if (value is string) cell.NumberFormat = "@"; cell.Value2 = value; }
        }
    }
    private static void PickerChecks(object application, string source)
    {
        string hash = DataMigrationHost.Digest(source);
        using (var host = new DataMigrationHost(application))
        {
            Check(host.PickSourceRange(source, "源表", "", s => ((dynamic)s).Range["B1:C3"]) == "B1:C3", "WPS source picker validates native range identity and returns coordinates");
        }
        using (var host = new DataMigrationHost(application)) Check(host.PickSourceRange(source, "源表", "B1:C3", s => false) == null, "WPS source picker cancellation returns no replacement");
        Check(DataMigrationHost.Digest(source) == hash, "source picker preserves original file bytes");
        Check(!Directory.GetFiles(Path.GetDirectoryName(source), ".xy-migration-read-*").Any(), "source picker and cancellation remove their disposable read copies");
    }
    [STAThread] private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        using (var session = new ProbeSession("data-migration-live"))
        {
            int code = 1;
            try
            {
                object application = session.Take(ProbeSession.Connect(new[] { "KET.Application", "et.Application" }, false));
                if (application == null) { Console.WriteLine("SKIP: WPS 表格 COM 不可用。"); return session.Finish(2); }
                dynamic app = application;
                Console.WriteLine("HOST: " + app.Name + " " + app.Version + " PATH=" + app.Path);
                object originalBook = null, originalSheet = null;
                try { originalBook = session.Take(app.ActiveWorkbook); originalSheet = session.Take(app.ActiveSheet); } catch { }
                session.OnClose(delegate { if (originalBook != null) ((dynamic)originalBook).Activate(); if (originalSheet != null) ((dynamic)originalSheet).Activate(); });
                object visible = app.Visible;
                session.Override(() => (object)app.Visible, value => app.Visible = (dynamic)value, visible, "Visible");
                session.Override(() => (object)app.DisplayAlerts, value => app.DisplayAlerts = (dynamic)value, false, "DisplayAlerts");
                session.Override(() => (object)app.ScreenUpdating, value => app.ScreenUpdating = (dynamic)value, false, "ScreenUpdating");
                session.Override(() => (object)app.EnableEvents, value => app.EnableEvents = (dynamic)value, false, "EnableEvents");
                if ((args.Length==4&&(args[0]=="--repro"||args[0]=="--diagnose"))||(args.Length==5&&args[0]=="--repair")) { Reproduce(session, application, args); return session.Finish(0); }
                string third = session.Output("第三方.xlsx"), source = session.Output("源.xlsx");
                Make(session, application, third, 51, "明细", delegate(object book, object sheet) { Cell(session, sheet, "A1", 37.0, false); });
                Make(session, application, source, 51, "源表", delegate(object book, object sheet)
                {
                    Cell(session, sheet, "A1", 10.0, false);
                    string external = "'" + Path.Combine(Path.GetDirectoryName(third), "[" + Path.GetFileName(third) + "]明细").Replace("'", "''") + "'!A1";
                    Cell(session, sheet, "B1", "=A1+$A$1+" + external + "+Rate", true);
                    Cell(session, sheet, "C1", "=普通文本", false); Cell(session, sheet, "B2", 123.5, false);
                    using (var refs = session.References())
                    {
                        dynamic names = refs.Take(((dynamic)book).Names); refs.Take(names.Add("Rate", "='源表'!$A$1", true));
                        dynamic merge = refs.Take(((dynamic)sheet).Range["B3:C3"]); merge.Merge(); merge.Value2 = "合并标题";
                        dynamic font = refs.Take(merge.Font); font.Bold = true; font.Color = 255;
                        dynamic fill = refs.Take(merge.Interior); fill.Color = 15658734;
                    }
                });
                PickerChecks(application, source);
                if (args.Length == 1 && args[0] == "--picker") return session.Finish(0);
                if (args.Length == 1 && args[0] == "--picker-ui")
                {
                    session.Override(() => (object)app.Visible, value => app.Visible = (dynamic)value, true, "picker visibility");
                    session.Override(() => (object)app.ScreenUpdating, value => app.ScreenUpdating = (dynamic)value, true, "picker redraw");
                    string originalHash = DataMigrationHost.Digest(source);
                    using (var host = new DataMigrationHost(application))
                    {
                        string picked = host.PickSourceRange(source, "源表", "", null);
                        Console.WriteLine("MOUSE_SELECTION: " + picked);
                        Check(picked == "B1:C3", "native mouse drag and OK return B1:C3");
                        Check(host.PickSourceRange(source, "源表", picked, null) == null, "native cancel keeps the existing address available");
                    }
                    Check(DataMigrationHost.Digest(source) == originalHash && !Directory.GetFiles(session.DirectoryPath, ".xy-migration-read-*").Any(), "native mouse selection and cancel leave originals intact and remove own copies");
                    return session.Finish(0);
                }
                var targets = new List<MigrationTarget>();
                foreach (var pair in new[] { Tuple.Create(".xlsx", 51), Tuple.Create(".xlsm", 52), Tuple.Create(".xls", 56), Tuple.Create(".et", -4143) })
                {
                    string target = session.Output("目标" + pair.Item1);
                    Make(session, application, target, pair.Item2, "目标表", delegate(object book, object sheet)
                    {
                        Cell(session, sheet, "A1", 20.0, false); Cell(session, sheet, "A100", 30.0, false); Cell(session, sheet, "B100", "原内容", false);
                        using (var refs = session.References()) { dynamic cell = refs.Take(((dynamic)sheet).Range["B100"]); cell.ColumnWidth = 19; cell.RowHeight = 24; }
                    });
                    targets.Add(new MigrationTarget { Path = target, Sheet = "目标表", Start = "B100" });
                }
                var hashes = targets.Select(t => t.Path).Concat(new[] { source, third }).ToDictionary(p => p, DataMigrationHost.Digest);
                DataMigrationHost.DiagnosticRootOverride = Path.Combine(session.DirectoryPath, "diagnostics");
                using (var host = new DataMigrationHost(application))
                {
                    Drain(host.Preview(new MigrationRequest { SourcePath = source, SourceSheet = "源表", SourceRange = "B1:C3", Targets = targets }));
                    foreach (var target in targets) Check(target.Ready, "预检 " + Path.GetExtension(target.Path) + ": " + target.Error);
                    Drain(host.Execute());
                    foreach (var target in targets)
                    {
                        Check(target.Status == "成功" && File.Exists(target.Output), "迁移并重开校验 " + Path.GetExtension(target.Path) + ": " + target.Error);
                        session.TrackFile(target.Output);
                    }
                }
                foreach (var pair in hashes) Check(DataMigrationHost.Digest(pair.Key) == pair.Value, "原文件摘要不变：" + Path.GetFileName(pair.Key));
                if (args.Length == 1)
                {
                    string imported = session.Output("生成器测试源.xlsx"), targetPath = session.Output("生成器目标.xlsx");
                    File.Copy(Path.GetFullPath(args[0]), imported, false);
                    string importedHash = DataMigrationHost.Digest(imported); int bookCount;
                    using (var refs = session.References()) { dynamic books = refs.Take(app.Workbooks); bookCount = Convert.ToInt32(books.Count); }
                    using (var host = new DataMigrationHost(application)) Check(host.Sheets(imported).SequenceEqual(new[] { "迁移源", "辅助明细" }), "直接读取生成器测试源的表名");
                    using (var refs = session.References()) { dynamic books = refs.Take(app.Workbooks); Check(Convert.ToInt32(books.Count) == bookCount, "选择源文件不打开或关闭宿主工作簿"); }
                    var generatedTargets = new List<MigrationTarget>(); var generatedHashes = new Dictionary<string, string>();
                    foreach (string generatedPath in new[] { targetPath, session.Output("生成器目标 (2).xlsx") })
                    {
                        Make(session, application, generatedPath, 51, "目标表", delegate(object book, object sheet)
                        {
                            Cell(session, sheet, "C6", 10.0, false);
                            using (var refs = session.References()) { dynamic sheets = refs.Take(((dynamic)book).Worksheets); object auxiliary = refs.Take(sheets.Add()); ((dynamic)auxiliary).Name = "辅助明细"; Cell(session, auxiliary, "B2", 100.0, false); }
                        });
                        generatedHashes[generatedPath] = DataMigrationHost.Digest(generatedPath);
                        generatedTargets.Add(new MigrationTarget { Path = generatedPath, Sheet = "目标表", Start = "B100" });
                    }
                    session.Override(() => (object)app.EnableEvents, value => app.EnableEvents = (dynamic)value, true, "generated source form events");
                    using (var host = new DataMigrationHost(application))
                    {
                        RunFormBatch(host, new MigrationRequest { SourcePath = imported, SourceSheet = "迁移源", SourceRange = "B5:I22", Targets = generatedTargets });
                        foreach (var target in generatedTargets)
                        {
                            Check(target.Status == "成功", "正式窗口消息泵完成两个目标的 144 格迁移与重开校验：" + target.Error);
                            if (target.Output.Length > 0) session.TrackFile(target.Output);
                            using (var refs = session.References())
                            {
                                dynamic books = refs.Take(app.Workbooks); dynamic result = refs.Take(books.Open(target.Output, 0, true));
                                try
                                {
                                    dynamic outputSheet = refs.Take(result.ActiveSheet), header = refs.Take(outputSheet.Range["B100"]), formula = refs.Take(outputSheet.Range["D101"]), date = refs.Take(outputSheet.Range["D113"]);
                                    Check(Convert.ToString(outputSheet.Name) == "目标表" && Convert.ToString(header.Value2) == "测试项目" && Convert.ToString(formula.Formula) == "=C101", "重新打开结果显示目标表且包含标题和实际公式");
                                    Check(Convert.ToDouble(date.Value2) == new DateTime(2026, 10, 29).ToOADate() && Convert.ToString(date.Text) == "2026-10-29", "ISO 日期迁移后仍为正确日历日期而非 1905 年");
                                    Check(Path.GetFileName(target.Output).StartsWith(Path.GetFileNameWithoutExtension(target.Path) + "_迁移结果", StringComparison.Ordinal), "中文正式结果文件名保留目标原名");
                                    dynamic windows = refs.Take(result.Windows), window = refs.Take(windows.Item(1));
                                    Check(Convert.ToInt32(window.ScrollRow) == 100, "重新打开结果直接显示迁移位置而非空白 A1");
                                }
                                finally { result.Close(false); }
                            }
                        }
                        Check(generatedTargets.Select(t => t.Output).Distinct().Count() == 2, "两个目标生成两个独立结果副本");
                    }
                    Check(DataMigrationHost.Digest(imported) == importedHash && generatedHashes.All(p => DataMigrationHost.Digest(p.Key) == p.Value), "生成器样本原文件摘要不变");
                    var circular = new MigrationTarget { Path = targetPath, Sheet = "目标表", Start = "A1" };
                    using (var host = new DataMigrationHost(application))
                    {
                        Drain(host.Preview(new MigrationRequest { SourcePath = imported, SourceSheet = "迁移源", SourceRange = "B5:I22", Targets = new List<MigrationTarget> { circular } }));
                        Check(!circular.Ready && circular.Error.Contains("C4") && circular.Error.Contains("自身") && circular.Output.Length == 0, "A1 起点引入直接自引用时预检明确阻止，未生成结果");
                    }
                    string reversedPath = session.Output("逆序范围目标.xlsx");
                    Make(session, application, reversedPath, 51, "目标表", delegate(object book, object sheet) { Cell(session, sheet, "C6", 100.0, false); Cell(session, sheet, "B5", 40.0, false); });
                    var reversed = new MigrationTarget { Path = reversedPath, Sheet = "目标表", Start = "C17" };
                    using (var host = new DataMigrationHost(application))
                    {
                        RunFormBatch(host, new MigrationRequest { SourcePath = imported, SourceSheet = "迁移源", SourceRange = "D21", Targets = new List<MigrationTarget> { reversed } });
                        Check(reversed.Status == "成功", "WPS 将逆序范围改为 B5:$C$6 后仍通过等价核验：" + reversed.Error);
                        if (reversed.Output.Length > 0) session.TrackFile(reversed.Output);
                    }
                }
                using (var form = new DataMigrationForm()) Check(!form.Visible, "真实宿主存在时构造窗口不自动运行迁移");
                // Every file under this unique fixture directory was generated by
                // this session; TrackFile still enforces the cleanup boundary.
                foreach (string file in Directory.GetFiles(session.DirectoryPath, "*", SearchOption.AllDirectories)) session.TrackFile(file);
                Console.WriteLine("PASS: four WPS file formats; UI focus and mouse interaction still require manual acceptance."); code = 0;
            }
            catch (Exception error) { Console.WriteLine("FAIL: " + error); }
            finally { DataMigrationHost.DiagnosticRootOverride = null; }
            return session.Finish(code);
        }
    }
}
