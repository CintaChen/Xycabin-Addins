using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

// Exercises the actual production reader and reference scope through reflection.
// All Application/Selection/ROT references below are managed fake objects.
public static class EtSelectionReaderProbe
{
    public class Reference
    {
        public readonly string Role;
        public int Outstanding, Acquisitions, Releases;
        public bool FailRelease;
        public Reference(string name) { Role = name; }
        public object Acquire() { Outstanding++; Acquisitions++; return this; }
    }
    public sealed class Book : Reference
    {
        public bool FailName;
        public Book() : base("book") { }
        public string FullName { get { if (FailName) throw new InvalidOperationException("injected FullName"); return "C:\\sample\\source.xlsx"; } }
    }
    public sealed class Sheet : Reference
    {
        public readonly Book Book;
        public Sheet(Book book) : base("sheet") { Book = book; }
        public string Name { get { return "Sheet1"; } }
        public object Parent { get { return Book.Acquire(); } }
    }
    public sealed class Area : Reference
    {
        public string Address { get; set; }
        public readonly Sheet Sheet;
        public Area(string name, string address, Sheet sheet) : base(name) { Address = address; Sheet = sheet; }
        public object Worksheet { get { return Sheet.Acquire(); } }
    }
    public sealed class Areas : Reference
    {
        public readonly List<Area> Items;
        public bool FailCount;
        public int FailItem;
        public Areas(List<Area> items) : base("areas") { Items = items; }
        public int Count { get { if (FailCount) throw new InvalidOperationException("injected Count"); return Items.Count; } }
        public object Item(int index) { if (FailItem == index) throw new InvalidOperationException("injected Item"); return Items[index - 1].Acquire(); }
    }
    public sealed class Selection : Reference
    {
        public string Address { get; set; }
        public readonly Areas Items;
        public Selection(Areas items) : base("selection") { Items = items; Address = "$H$1:$H$3"; }
        public object Areas { get { return Items.Acquire(); } }
    }
    public sealed class Host : Reference
    {
        public readonly Selection Selected;
        public bool FailSelection;
        public int QuitCalls;
        public Host(Selection selected) : base("borrowed-host") { Selected = selected; Outstanding = 1; }
        public object Selection { get { if (FailSelection) throw new InvalidOperationException("injected Selection"); return Selected.Acquire(); } }
        public void Quit() { QuitCalls++; }
    }
    private sealed class Sample
    {
        internal readonly Book Book;
        internal readonly Sheet Sheet;
        internal readonly Area First, Second;
        internal readonly Areas Areas;
        internal readonly Selection Selection;
        internal readonly Host Host;
        internal readonly Reference[] All;
        internal Sample()
        {
            Book = new Book(); Sheet = new Sheet(Book);
            First = new Area("first", "$H$1", Sheet); Second = new Area("second", "$H$3", Sheet);
            Areas = new Areas(new List<Area> { First, Second }); Selection = new Selection(Areas); Host = new Host(Selection);
            All = new Reference[] { Book, Sheet, First, Second, Areas, Selection };
        }
    }

    private static Type readerType, scopeType;
    private static MethodInfo own, add, read;
    private static int checks;

    private static void Check(bool value, string name)
    {
        if (!value) throw new InvalidOperationException(name);
        checks++;
        Console.WriteLine("PASS: " + name);
    }

    private static object NewScope(List<string> released, List<string> diagnostics)
    {
        Func<object, bool> isCom = value => value is Reference;
        Action<object> release = value =>
        {
            Reference reference = (Reference)value;
            released.Add(reference.Role);
            reference.Releases++;
            if (reference.FailRelease) throw new InvalidOperationException("release " + reference.Role);
            reference.Outstanding--;
            if (reference.Outstanding < 0) throw new InvalidOperationException("over-release " + reference.Role);
        };
        return Activator.CreateInstance(scopeType, BindingFlags.Instance | BindingFlags.NonPublic, null,
            new object[] { isCom, release, new Action<string>(diagnostics.Add) }, null);
    }

    private static object Read(Sample sample, List<string> released, List<string> diagnostics, out string note)
    {
        object[] args = { sample.Host, null, NewScope(released, diagnostics) };
        object result = read.Invoke(null, args);
        note = (string)args[1];
        return result;
    }

    private static void Balanced(Sample sample, string name)
    {
        Check(sample.All.All(reference => reference.Outstanding == 0 && reference.Releases == reference.Acquisitions), name);
        Check(sample.Host.Outstanding == 1 && sample.Host.Releases == 0 && sample.Host.QuitCalls == 0, "borrowed Application survives " + name);
    }

    private static void DuplicateDiscovery()
    {
        Sample sample = new Sample();
        List<object> candidates = new List<object>();
        List<string> released = new List<string>(), diagnostics = new List<string>();
        object scope = NewScope(released, diagnostics);
        for (int i = 0; i < 3; i++) add.Invoke(null, new object[] { candidates, sample.Host.Acquire(), scope });
        Check(candidates.Count == 1 && sample.Host.Outstanding == 4, "ROT/ProgID aliases process one identity but retain every acquisition");
        ((IDisposable)scope).Dispose();
        Check(sample.Host.Releases == 3 && sample.Host.Outstanding == 1 && diagnostics.Count == 0, "release duplicate acquisitions once each and preserve caller reference");
        ((IDisposable)scope).Dispose();
        Check(sample.Host.Releases == 3, "reference cleanup is idempotent");
    }

    private static void SelectionSuccessAndEarlyReturns()
    {
        Sample sample = new Sample();
        List<string> released = new List<string>(), diagnostics = new List<string>();
        string note;
        object result = Read(sample, released, diagnostics, out note);
        Check(result != null && note == null, "actual reader accepts disjoint H1/H3 areas");
        IList areas = (IList)result.GetType().GetField("Areas").GetValue(result);
        Check(areas.Count == 2 && (string)result.GetType().GetField("WorkbookPath").GetValue(result) == sample.Book.FullName,
            "actual reader preserves source workbook and two area mappings");
        Check((string)result.GetType().GetField("SheetName").GetValue(result) == "Sheet1", "actual reader preserves source sheet");
        Check(released.Count == 8 && sample.Book.Releases == 2 && sample.Sheet.Releases == 2, "release repeated Worksheet/Parent acquisitions per area");
        Balanced(sample, "successful selection");
        Check(diagnostics.Count == 0, "successful reference cleanup has no diagnostics");

        sample = new Sample(); sample.Host.FailSelection = true; released.Clear();
        Check(Read(sample, released, diagnostics, out note) == null && note.Contains("无可用选区") && released.Count == 0, "Selection failure has no unowned reference cleanup");
        Balanced(sample, "Selection exception");

        sample = new Sample(); sample.Selection.Address = null; released.Clear();
        Check(Read(sample, released, diagnostics, out note) == null && note.Contains("不是单元格") && released.Count == 1, "non-cell selection early return releases Selection");
        Balanced(sample, "selection address rejection");

        sample = new Sample(); sample.Areas.FailCount = true; released.Clear();
        Check(Read(sample, released, diagnostics, out note) == null && note.Contains("只有一片") && released.Count == 2, "Areas.Count exception releases both acquired references");
        Balanced(sample, "area count exception");

        sample = new Sample(); sample.Areas.FailItem = 2; released.Clear();
        Check(Read(sample, released, diagnostics, out note) == null && note.Contains("子区域 2") && released.Count == 5,
            "later Item exception releases earlier area, Worksheet, Parent and selection");
        Balanced(sample, "later area exception");

        sample = new Sample(); sample.Second.Address = null; released.Clear();
        Check(Read(sample, released, diagnostics, out note) == null && note.Contains("无地址") && released.Count == 8, "malformed second area releases all acquired metadata");
        Balanced(sample, "malformed area");

        sample = new Sample(); sample.Book.FailName = true; released.Clear();
        result = Read(sample, released, diagnostics, out note);
        Check(result != null && result.GetType().GetField("WorkbookPath").GetValue(result) == null,
            "unreadable FullName remains an unverified candidate for caller validation");
        Balanced(sample, "FullName exception");

        sample = new Sample(); sample.First.Address = "invalid-address"; released.Clear();
        Check(Read(sample, released, diagnostics, out note) == null && note.Contains("无法解析"), "invalid area parser path fails without retaining COM metadata");
        Balanced(sample, "parser rejection");
    }

    private static void ScopeFinallyAndReleaseFailures()
    {
        List<string> released = new List<string>(), diagnostics = new List<string>();
        Reference rot = new Reference("ROT"), enumerator = new Reference("enumerator");
        Reference moniker = new Reference("moniker"), context = new Reference("bind-context");
        try
        {
            using ((IDisposable)NewScope(released, diagnostics)) { } // Empty native return.
            object outer = NewScope(released, diagnostics);
            using ((IDisposable)outer)
            {
                own.Invoke(outer, new[] { rot.Acquire() });
                own.Invoke(outer, new[] { enumerator.Acquire() });
                object inner = NewScope(released, diagnostics);
                using ((IDisposable)inner)
                {
                    own.Invoke(inner, new[] { moniker.Acquire() });
                    own.Invoke(inner, new[] { context.Acquire() });
                    throw new InvalidOperationException("injected native enumeration failure");
                }
            }
        }
        catch (InvalidOperationException) { }
        Check(string.Join(",", released) == "bind-context,moniker,enumerator,ROT", "production scope finally releases ROT, enumerator, moniker and context in reverse order");
        Check(new[] { rot, enumerator, moniker, context }.All(reference => reference.Outstanding == 0), "native enumeration exception leaves no acquired references");

        Sample sample = new Sample();
        sample.Second.FailRelease = true;
        released.Clear(); diagnostics.Clear();
        string note;
        Check(Read(sample, released, diagnostics, out note) != null, "metadata reads remain available when a cleanup callback fails");
        Check(released.Count == 8 && diagnostics.Count == 1 && diagnostics[0].Contains("second"), "release failure is diagnosed and remaining references are still released");
        Check(sample.All.Where(reference => reference != sample.Second).All(reference => reference.Outstanding == 0), "release failure does not leak unrelated COM references");
        Check(sample.Host.Outstanding == 1 && sample.Host.Releases == 0, "release failure never force-releases the borrowed Application");
    }

    private static int Main()
    {
        try
        {
            Assembly assembly = Assembly.LoadFrom(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WordUpdater.dll"));
            readerType = assembly.GetType("WordUpdater.EtSelectionReader", true);
            scopeType = assembly.GetType("WordUpdater.EtReferenceScope", true);
            own = scopeType.GetMethod("Own", BindingFlags.Instance | BindingFlags.NonPublic);
            add = readerType.GetMethod("AddApp", BindingFlags.Static | BindingFlags.NonPublic);
            read = readerType.GetMethods(BindingFlags.Static | BindingFlags.NonPublic)
                .Single(method => method.Name == "ReadFrom" && method.GetParameters().Length == 3);
            DuplicateDiscovery();
            SelectionSuccessAndEarlyReturns();
            ScopeFinallyAndReleaseFailures();
            Console.WriteLine("EtSelectionReaderProbe: " + checks + " PASS, 0 FAIL; fake references only, no WPS attached");
            return 0;
        }
        catch (Exception error) { Console.WriteLine("EtSelectionReaderProbe FAIL: " + error); return 1; }
    }
}
