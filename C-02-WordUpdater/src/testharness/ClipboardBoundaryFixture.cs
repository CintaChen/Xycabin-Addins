using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

// Reflection invokes the compiled probe and actual production geometry/engine.
// All host references are managed fakes; no COM activation or desktop input.
public static class ClipboardBoundaryFixture
{
    public sealed class Numbers { public int Count { get; set; } }
    public sealed class Range
    {
        public int Start { get; set; }
        public int End { get; set; }
        public object Cells { get; set; }
        public Tables Tables { get; set; }
        public Numbers Fields { get; set; }
        public Document Document;
        public string WordOpenXML;
        public bool Content;
        private string text = "";
        public string Text
        {
            get { return Document == null ? text : Content ? Document.Buffer : Document.Buffer.Substring(Start, End - Start); }
            set { if (Document == null) text = value; else { Document.Buffer = value; End = value.Length; } }
        }
    }
    public sealed class Cell
    {
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public Range Range { get; set; }
    }
    public sealed class Cells
    {
        public readonly List<Cell> Values = new List<Cell>();
        public int Count { get { return Values.Count; } }
        public Cell Item(int index) { return Values[index - 1]; }
    }
    public sealed class Table
    {
        public bool Uniform { get { return true; } }
        public Numbers Rows { get; set; }
        public Numbers Columns { get; set; }
        public Range Range { get; set; }
    }
    public sealed class Tables
    {
        public readonly List<Table> Values = new List<Table>();
        public Document Document;
        public int ThrowOnAdd, Adds;
        public bool Merge;
        public readonly List<int> Insertions = new List<int>();
        public int Count { get { return Values.Count; } }
        public Table Item(int index) { return Values[index - 1]; }
        public object Add(Range range, int rows, int columns)
        {
            Adds++; Insertions.Add(range.Start);
            if (Adds == ThrowOnAdd) throw new COMException("injected table insertion", unchecked((int)0x80004005));
            if (Merge && Values.Count > 0) return Values[0];
            int length = rows * columns + rows;
            Document.Buffer = Document.Buffer.Insert(range.Start, new string('#', length));
            foreach (Table previous in Values)
                if (previous.Range.Start >= range.Start) { previous.Range.Start += length; previous.Range.End += length; }
            var table = new Table { Rows = new Numbers { Count = rows }, Columns = new Numbers { Count = columns },
                Range = new Range { Start = range.Start, End = range.Start + length, Document = Document } };
            Values.Add(table); return table;
        }
    }
    public sealed class Document
    {
        public string Buffer = "sample\r";
        public readonly Tables Tables;
        public bool BadAnchor;
        public Document() { Tables = new Tables { Document = this }; }
        public Range Content { get { return new Range { Start = 0, End = Buffer.Length, Document = this, Content = true, Fields = new Numbers() }; } }
        public Range Range(int start, int end)
        {
            if (BadAnchor && end > start) return new Range { Start = start, End = end, Text = "wrong anchor" };
            if (start < 0 || end < start || end > Buffer.Length) throw new ArgumentOutOfRangeException("range");
            return new Range { Start = start, End = end, Document = this };
        }
    }
    public sealed class Selection
    {
        public Range Range { get; set; }
        public Tables Tables { get; set; }
        public object CellValues;
        public bool ThrowCells, InTable;
        public Action OnInformation;
        public object Cells { get { if (ThrowCells) throw new COMException("injected Selection.Cells"); return CellValues; } }
        public bool Information(int value) { if (OnInformation != null) OnInformation(); return InTable; }
    }
    public sealed class Host
    {
        public Selection Selection { get; set; }
        public Numbers Documents { get; set; }
    }

    static Type sessionType, probeType;
    static int checks;
    static void Check(bool condition, string label) { if (!condition) throw new Exception(label); checks++; }
    static object Invoke(string name, params object[] args)
    {
        try { return probeType.GetMethod(name, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, args); }
        catch (TargetInvocationException error) { throw error.InnerException; }
    }
    static void Throws(Action action, string fragment)
    {
        try { action(); }
        catch (Exception error) { Check(error.Message.Contains(fragment), "expected failure: " + fragment); return; }
        throw new Exception("expected failure: " + fragment);
    }
    static object NewSession()
    {
        foreach (ConstructorInfo ctor in sessionType.GetConstructors(BindingFlags.NonPublic | BindingFlags.Instance))
            if (ctor.GetParameters().Length == 5)
                return ctor.Invoke(new object[] { "clipboard-boundary-fixture", Path.GetTempPath(), new Action<string>(text => { }),
                    new Func<object, bool>(value => false), new Action<object>(value => { }) });
        throw new Exception("missing injectable session");
    }
    static int Finish(object session, int code) { return (int)sessionType.GetMethod("Finish", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(session, new object[] { code }); }
    static Table TableSample()
    {
        string row = "<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>";
        return new Table { Rows = new Numbers { Count = 3 }, Columns = new Numbers { Count = 3 },
            Range = new Range { Start = 0, End = 12, WordOpenXML = "<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + row + row + row + "</w:tbl>" } };
    }
    static Selection Outside(Table table)
    {
        var tables = new Tables(); tables.Values.Add(table);
        return new Selection { Range = new Range { Start = 10, End = 20 }, Tables = tables };
    }
    static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        try
        {
            string executable = args.Length == 1 ? Path.GetFullPath(args[0]) : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ClipboardAreaProbe.exe");
            Assembly assembly = Assembly.LoadFrom(executable);
            sessionType = assembly.GetType("ProbeSession", true); probeType = assembly.GetType("ClipboardAreaProbe", true);
            object session = NewSession();
            try
            {
                Table table = TableSample(); Selection outside = Outside(table);
                Check(Invoke("ReadRect", session, "null outside cells", outside, table).ToString() == "OutOfTable", "outside null cells rejected before diagnostics");
                outside.ThrowCells = true;
                Check(Invoke("ReadRect", session, "throwing outside cells", outside, table).ToString() == "OutOfTable", "outside Cells COM failure does not mask geometry");
                var cells = new Cells(); var parents = new Tables(); parents.Values.Add(table);
                cells.Values.Add(new Cell { RowIndex = 2, ColumnIndex = 2, Range = new Range { Start = 5, End = 6, Tables = parents } });
                var fallback = new Selection { Range = new Range { Start = 5, End = 5, Cells = cells } };
                Check(Invoke("ReadRect", session, "range cells fallback", fallback, table).ToString() == "Ok", "positive selection falls back to Range.Cells");
                var missing = new Selection { Range = new Range { Start = 5, End = 5 } };
                object unreadable = Invoke("ReadRect", session, "missing positive cells", missing, table);
                Check(unreadable.ToString() == "ReadError", "missing in-table cells remain ReadError");
                Array expected = Array.CreateInstance(unreadable.GetType(), 1); expected.SetValue(Enum.Parse(unreadable.GetType(), "Ok"), 0);
                Throws(() => Invoke("RequireStatus", unreadable, expected), "Unexpected target geometry");

                Document document = new Document();
                var created = (object[])Invoke("CreateBoundaryTables", session, document);
                Check(created.Length == 2 && document.Tables.Count == 2, "two independent tables created");
                Check(document.Tables.Insertions[0] > 0 && document.Tables.Insertions[1] == 0, "second table uses interior text anchor before first table");
                Check(((Table)created[1]).Range.Start > ((Table)created[0]).Range.End, "paragraph gap survives first-table insertion");
                var badAnchor = new Document { BadAnchor = true };
                Throws(() => Invoke("CreateBoundaryTables", session, badAnchor), "text anchor");
                Check(badAnchor.Tables.Adds == 0, "unverified anchor never inserts a table");
                var merged = new Document(); merged.Tables.Merge = true;
                Throws(() => Invoke("CreateBoundaryTables", session, merged), "two separate");
                var failedAdd = new Document(); failedAdd.Tables.ThrowOnAdd = 1;
                Throws(() => Invoke("CreateBoundaryTables", session, failedAdd), "injected table insertion");

                var host = new Host { Selection = Outside(table), Documents = new Numbers { Count = 1 } };
                Invoke("ProbeRejection", session, "outside engine rejection", host, new Document(), table);
                Check(true, "actual engine rejects outside selection without writes");
                host.Selection.InTable = true;
                Invoke("ProbeRejection", session, "cross table engine rejection", host, new Document(), table);
                Check(true, "actual engine returns SelectionOutOfTable before clipboard or field writes");
                var changed = new Document(); host.Selection.InTable = false; host.Selection.OnInformation = () => changed.Buffer += "mutated";
                Throws(() => Invoke("ProbeRejection", session, "changed content", host, changed, table), "changed document content");
                host.Selection.OnInformation = null; host.Documents.Count = 0;
                Throws(() => Invoke("ProbeRejection", session, "wrong engine failure", host, new Document(), table), "not explicitly rejected");
                Check(Finish(session, 0) == 0, "successful fake-session cleanup returns zero");
            }
            finally { ((IDisposable)session).Dispose(); }
            object failing = NewSession(); string directory = (string)sessionType.GetField("DirectoryPath", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(failing);
            try
            {
                Invoke("Observe", failing, "injected construction failure", new Action(() => { throw new COMException("injected E_FAIL"); }));
                Check(Finish(failing, 0) == 1, "observation failure cannot be overwritten by success");
                Check(Directory.Exists(directory), "failed observation retains evidence directory");
            }
            finally { ((IDisposable)failing).Dispose(); if (Directory.Exists(directory) && Directory.GetFileSystemEntries(directory).Length == 0) Directory.Delete(directory); }
            Console.WriteLine("ClipboardBoundaryFixture: " + checks + " PASS, 0 FAIL; fake hosts only, no WPS attached");
            return 0;
        }
        catch (Exception error) { Console.WriteLine(error); return 1; }
    }
}
