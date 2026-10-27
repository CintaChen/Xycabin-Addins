using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

// Public mock objects exercise save/identity guards across the assembly boundary.
public sealed class PdfSaveProbeDocument
{
    public string FullName; public bool ReadOnly, Saved; public int Saves, Exports;
    public void Save() { Saves++; }
    public void ExportAsFixedFormat(params object[] args) { Exports++; }
}
public sealed class PdfSaveProbeApplication
{
    public object ActiveDocument, ActiveWorkbook, ActiveSheet, Selection;
}
internal static class PdfCurrentDocumentProbe
{
    private const BindingFlags All = BindingFlags.Static | BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint process);
    private static Assembly assembly; private static string ns; private static int checks;
    private static void Check(bool value, string message) { if (!value) throw new Exception(message); checks++; }
    private static object Request(string scope, string output)
    {
        object request = Activator.CreateInstance(assembly.GetType(ns + ".PdfToolRequest"), true);
        Set(request, "Scope", scope); Set(request, "Output", output); Set(request, "Ranges", ""); return request;
    }
    private static void Set(object obj, string name, object value) { obj.GetType().GetField(name, All).SetValue(obj, value); }
    private static object Session(object app) { return Activator.CreateInstance(assembly.GetType(ns + ".PdfCurrentDocument"), All, null, new[] { app }, null); }
    private static string Export(object session, object request) { return (string)session.GetType().GetMethod("Export", All).Invoke(session, new[] { request }); }
    private static void ExpectFailure(object session, object request)
    { try { Export(session, request); throw new Exception("Expected export refusal"); } catch (TargetInvocationException) { checks++; } }
    private static void Offline(string output)
    {
        string input = Path.Combine(output, "save-guard.txt"); File.WriteAllText(input, "original");
        var doc = new PdfSaveProbeDocument { FullName = input, ReadOnly = true };
        var app = new PdfSaveProbeApplication { ActiveDocument = doc, ActiveWorkbook = doc };
        object request = Request("book", output);
        using (var session = (IDisposable)Session(app)) { ExpectFailure(session, request); Check(doc.Saves == 0 && doc.Exports == 0, "read-only source was touched"); }
        doc.ReadOnly = false;
        using (var session = (IDisposable)Session(app)) { ExpectFailure(session, request); Check(doc.Saves == 1 && doc.Exports == 0, "unsaved cancellation reached export"); }
        using (var session = (IDisposable)Session(app))
        { app.ActiveDocument = app.ActiveWorkbook = new PdfSaveProbeDocument(); ExpectFailure(session, request); Check(doc.Saves == 1, "changed active document was saved"); }
        app.ActiveDocument = app.ActiveWorkbook = doc;
        using (var session = (IDisposable)Session(app))
        { doc.FullName = input + ".renamed"; ExpectFailure(session, request); Check(doc.Saves == 1, "renamed source was saved"); }
        Check(WpsShared.SourcePaths.ReadAllText(input) == "original", "fixture was modified");
    }
    private static void Release(object obj) { if (obj != null && Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj); }
    private static void Live(string output)
    {
        bool word = ns == "WordUpdater"; string processName = word ? "WINWORD" : "EXCEL";
        var prior = Process.GetProcessesByName(processName);
        var priorIds = new HashSet<int>(); foreach (var process in prior) { priorIds.Add(process.Id); process.Dispose(); }
        object app = null, collection = null, document = null, first = null, second = null, range = null; bool owned = false;
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID(word ? "Word.Application" : "Excel.Application", true));
            collection = word ? (object)((dynamic)app).Documents : (object)((dynamic)app).Workbooks;
            if (Convert.ToInt32(((dynamic)collection).Count) != 0)
            { Console.WriteLine("SKIP live export: automation returned an instance with existing documents."); return; }
            document = ((dynamic)collection).Add();
            object ownWindow = null;
            try
            {
                ownWindow = word ? (object)((dynamic)document).ActiveWindow : (object)((dynamic)app).ActiveWindow;
                uint pid; GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)ownWindow).Hwnd)), out pid);
                owned = pid != 0 && !priorIds.Contains((int)pid);
            }
            finally { Release(ownWindow); }
            if (!owned) { Console.WriteLine("SKIP live export: no independently owned Office process."); return; }
            if (word)
            {
                range = ((dynamic)document).Content; ((dynamic)range).Text = "Current Word page one\fCurrent Word page two"; Release(range); range = null;
                ((dynamic)document).SaveAs2(Path.Combine(output, "current-word.docx"), 16);
                range = ((dynamic)document).Content; ((dynamic)range).InsertAfter(" UNSAVED-WORD-9988"); Release(range); range = null;
            }
            else
            {
                first = ((dynamic)document).Worksheets; second = ((dynamic)first).Item(1); Release(first); first = second; second = null;
                range = ((dynamic)first).Range["A:A"]; ((dynamic)range).ColumnWidth = 45; Release(range); range = null;
                range = ((dynamic)first).Range["A1"]; ((dynamic)range).Value2 = "SHEET-ONE"; Release(range); range = null;
                object sheets = ((dynamic)document).Worksheets;
                try { second = ((dynamic)sheets).Add(Type.Missing, first); } finally { Release(sheets); }
                range = ((dynamic)second).Range["A1"]; ((dynamic)range).Value2 = "SHEET-TWO"; Release(range); range = null;
                ((dynamic)document).SaveAs(Path.Combine(output, "current-excel.xlsx"), 51);
                ((dynamic)first).Activate(); range = ((dynamic)first).Range["A2"]; ((dynamic)range).Value2 = "UNSAVED-EXCEL-7788"; ((dynamic)range).Select(); Release(range); range = null;
            }
            using (var session = (IDisposable)Session(app))
            {
                string[] scopes = word ? new[] { "document" } : new[] { "book", "sheet", "selection" };
                foreach (string scope in scopes)
                {
                    object request = Request(scope, output); string scratch = Export(session, request);
                    try
                    {
                        Check(Convert.ToBoolean(((dynamic)document).Saved), "source changes were not saved");
                        string file = Path.Combine(scratch, "document.pdf"); Check(WpsShared.SourcePaths.Exists(file), "current PDF missing");
                        File.Copy(file, Path.Combine(output, scope + ".pdf"), true);
                        if (word)
                        {
                            Set(request, "Ranges", "2");
                            string report = (string)assembly.GetType(ns + ".PdfWorkerBridge").GetMethod("Run", All).Invoke(null, new object[] { request, new Func<bool>(() => false) });
                            Check(WpsShared.SourcePaths.Exists(Path.Combine(output, "current-word.pdf")), "Word page selection publication failed: " + report);
                        }
                    }
                    finally { foreach (string file in WpsShared.SourcePaths.GetFiles(scratch)) File.Delete(file); Directory.Delete(scratch); }
                }
            }
        }
        finally
        {
            Release(range); Release(first); Release(second);
            try { if (document != null) ((dynamic)document).Close(false); }
            finally { Release(document); Release(collection); try { if (app != null && owned) ((dynamic)app).Quit(); } finally { Release(app); } }
        }
    }
    [STAThread] private static int Main(string[] args)
    {
        try
        {
            assembly = Assembly.LoadFrom(Path.GetFullPath(args[0])); ns = assembly.GetType("WordUpdater.PdfToolsForm", false) != null ? "WordUpdater" : "Excel_Iterate";
            string output = Path.GetFullPath(args[1]); Directory.CreateDirectory(output); Offline(output);
            if (args.Length > 2 && args[2] == "--live") Live(output);
            Console.WriteLine("PASS: " + ns + " current-document guards/exports, " + checks + " checks"); return 0;
        }
        catch (Exception error) { Console.WriteLine(error); return 1; }
    }
}
