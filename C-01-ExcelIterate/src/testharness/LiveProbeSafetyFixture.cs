using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

// Offline regression of the exact lifecycle helper linked into all live probes.
// No ProgID lookup, COM activation, process enumeration, or WPS interaction.
internal static class LiveProbeSafetyFixture
{
    private sealed class Document
    {
        internal int CloseCalls;
        internal bool FailClose;
        internal void Close() { CloseCalls++; if (FailClose) throw new IOException("close failed"); }
    }

    private sealed class Host
    {
        internal bool Visible;
        internal object Alerts = false;
        internal int QuitCalls, KillCalls;
        internal readonly List<int> UserPids = new List<int> { 101 };
        internal readonly Document Existing = new Document();
        internal void Quit() { QuitCalls++; Existing.Close(); }
        internal void Kill(int pid) { KillCalls++; UserPids.Remove(pid); }
    }

    private static string root;
    private static int checks;
    private static readonly List<string> retained = new List<string>();

    private static void Check(bool passed, string description)
    {
        if (!passed) throw new InvalidOperationException(description);
        checks++;
        Console.WriteLine("PASS: " + description);
    }

    private static ProbeSession Session(string name, List<string> log, List<object> released)
    {
        ProbeSession session = new ProbeSession(name, root, log.Add, value => true, released.Add);
        retained.Add(session.DirectoryPath);
        return session;
    }

    private static void BorrowedAndSingleInstance()
    {
        Host host = new Host();
        List<string> log = new List<string>();
        List<object> released = new List<object>();
        int creates = 0;
        object attached = ProbeSession.Connect(new[] { "first", "second" }, false,
            id => id == "second" ? (object)host : null,
            id => { creates++; return host; }, log.Add);
        Check(ReferenceEquals(attached, host) && creates == 0, "attach searches all aliases before activation");
        ProbeSession session = Session("single-instance", log, released);
        object app = session.Take(ProbeSession.Connect(new[] { "host" }, false,
            id => { throw new COMException("not running"); },
            id => { creates++; return host; }, log.Add));
        Document probeDocument = new Document();
        session.Own(probeDocument, probeDocument.Close);
        object range = new object();
        using (ProbeReferences local = session.References()) local.Take(range);
        // A user starts another process while the probe is active. Ownership does
        // not depend on the before/after PID set, even if activation reused WPS.
        host.UserPids.Add(202);
        Check(ReferenceEquals(app, host) && creates == 1, "CreateInstance may return the existing host");
        Check(session.Finish(0) == 0, "borrowed host cleanup completes");
        Check(probeDocument.CloseCalls == 1 && host.Existing.CloseCalls == 0, "close only the created document");
        Check(host.QuitCalls == 0 && host.KillCalls == 0 && host.UserPids.Count == 2, "existing and concurrently opened user processes survive");
        Check(released.Count == 3 && ReferenceEquals(released[0], range)
            && ReferenceEquals(released[1], probeDocument) && ReferenceEquals(released[2], host), "release each acquired reference in reverse scope order");
        Check(!Directory.Exists(session.DirectoryPath), "successful empty probe directory removed");
        session.Dispose();
        Check(probeDocument.CloseCalls == 1 && released.Count == 3, "cleanup is idempotent");
    }

    private static void PropertySnapshots()
    {
        Host host = new Host();
        List<string> log = new List<string>();
        ProbeSession session = Session("properties", log, new List<object>());
        Check(session.Override(() => host.Visible, value => host.Visible = (bool)value, true, "Visible"), "known false Visible can be overridden");
        Check(session.Override(() => host.Alerts, value => host.Alerts = value, 0, "DisplayAlerts"), "known false DisplayAlerts can be overridden");
        Check(host.Visible && Equals(host.Alerts, 0), "temporary host properties are applied");
        int unknownWrites = 0;
        Check(!session.Override(() => { throw new COMException("unreadable"); }, value => unknownWrites++, true, "Unreadable"), "unreadable original remains untouched");
        Check(!session.Override(() => null, value => unknownWrites++, true, "Null"), "null original remains untouched");
        Check(!session.Override(() => Type.Missing, value => unknownWrites++, true, "Missing"), "missing original remains untouched");
        Check(unknownWrites == 0, "unknown properties are never written or restored");
        Check(session.Finish(0) == 0 && !host.Visible && Equals(host.Alerts, false), "restore exact original false values, not a default true");

        ProbeSession partial = Session("partial-setter", new List<string>(), new List<object>());
        bool valueNow = false;
        Check(!partial.Override(() => valueNow, value => {
            valueNow = (bool)value;
            if (valueNow) throw new COMException("setter mutated then failed");
        }, true, "Partial"), "failed setters may have mutated the host");
        Check(partial.Finish(0) == 0 && !valueNow, "restore is registered before a setter can throw");

        ProbeSession integer = Session("integer-alerts", new List<string>(), new List<object>());
        object alerts = 0;
        integer.Override(() => alerts, value => alerts = value, 2, "Word.DisplayAlerts");
        Check(integer.Finish(0) == 0 && alerts is int && Equals(alerts, 0), "restore original integer Word alert state");
    }

    private static void FailureAndEvidence()
    {
        List<string> log = new List<string>();
        List<object> released = new List<object>();
        ProbeSession session = Session("failed-run", log, released);
        string evidence = session.Output("source.xlsx");
        File.WriteAllText(evidence, "failure evidence");
        Document doc = new Document();
        session.Own(doc, doc.Close);
        object acquired = session.Take(new object());
        bool visible = false;
        session.Override(() => visible, value => visible = (bool)value, true, "Visible");
        session.Fail("injected business exception");
        Check(session.Finish(0) == 1, "later completion cannot overwrite an earlier failure");
        Check(doc.CloseCalls == 1 && !visible && released.Contains(acquired), "exceptions still close owned documents, restore properties and release references");
        Check(File.ReadAllText(evidence) == "failure evidence", "failure retains generated evidence");
        Check(log.Contains("PROBE_RESULT: FAIL"), "failure is explicitly reported");
        Check(session.Finish(2) == 1, "repeated finish cannot change a failed result");

        List<string> cleanupLog = new List<string>();
        int releases = 0, restorationAttempts = 0;
        ProbeSession badCleanup = new ProbeSession("cleanup-failure", root, cleanupLog.Add, item => true,
            item => { releases++; throw new COMException("release failed"); });
        retained.Add(badCleanup.DirectoryPath);
        string cleanupEvidence = badCleanup.Output("evidence.txt");
        File.WriteAllText(cleanupEvidence, "keep me");
        Document failingDoc = new Document { FailClose = true }, survivingDoc = new Document();
        badCleanup.Own(survivingDoc, survivingDoc.Close);
        badCleanup.Own(failingDoc, failingDoc.Close);
        badCleanup.Override(() => false, value => {
            restorationAttempts++;
            if (restorationAttempts > 1) throw new COMException("restore failed");
        }, true, "RestoreFailure");
        Check(badCleanup.Finish(0) == 1, "close/restore/release failures change PASS to FAIL");
        Check(survivingDoc.CloseCalls == 1 && failingDoc.CloseCalls == 1 && releases == 2 && restorationAttempts == 2, "cleanup failure does not suppress remaining finally work");
        Check(File.Exists(cleanupEvidence), "failed cleanup retains evidence");
    }

    private static void AttachOnlyAndExitCodes()
    {
        int creates = 0;
        List<string> log = new List<string>();
        object app = ProbeSession.Connect(new[] { "host1", "host2" }, true,
            id => { throw new COMException("not running"); },
            id => { creates++; return new Host(); }, log.Add);
        Check(app == null && creates == 0, "Inspect attach-only never activates a host");
        ProbeSession session = Session("skip", log, new List<object>());
        Check(session.Finish(2) == 2 && log.Contains("PROBE_RESULT: SKIP (not PASS)")
            && !log.Contains("PROBE_RESULT: PASS"), "SKIP returns 2 and is not reported as PASS");
        bool errorPropagated = false;
        try
        {
            ProbeSession.Connect(new[] { "host" }, false,
                id => { throw new InvalidOperationException("unexpected attach failure"); },
                id => new Host(), log.Add);
        }
        catch (InvalidOperationException) { errorPropagated = true; }
        Check(errorPropagated, "unexpected attach exceptions remain failures");
        ProbeSession invalid = Session("invalid-exit", new List<string>(), new List<object>());
        Check(invalid.Finish(99) == 1, "unexpected exit statuses normalize to FAIL");
    }

    private static void FileIsolation()
    {
        string caller = Path.Combine(root, "user-sentinel.xlsx");
        File.WriteAllText(caller, "USER CONTENT");
        ProbeSession session = Session("files", new List<string>(), new List<object>());
        bool refused = false;
        try { ProbeSession.RejectExistingOutput(caller); }
        catch (IOException) { refused = true; }
        Check(refused && File.ReadAllText(caller) == "USER CONTENT", "existing caller output is rejected without modification");
        string copied = session.CopyInput(caller, "copy.xlsx");
        File.WriteAllText(copied, "test writes to copy");
        Check(File.ReadAllText(caller) == "USER CONTENT", "repro writes only to a separate copy");
        refused = false;
        try { session.Export(copied, caller); } catch (IOException) { refused = true; }
        Check(refused && File.ReadAllText(caller) == "USER CONTENT", "export refuses existing output without overwriting");
        string export = Path.Combine(root, "new-export.xlsx");
        session.Export(copied, export);
        Check(File.ReadAllText(export) == "test writes to copy", "export preserves isolated source until copied");
        string concurrent = Path.Combine(session.DirectoryPath, "user-concurrent.txt");
        File.WriteAllText(concurrent, "do not delete");
        string concurrentEmptyDirectory = Path.Combine(session.DirectoryPath, "user-directory");
        Directory.CreateDirectory(concurrentEmptyDirectory);
        string backupDirectory = Path.Combine(session.DirectoryPath, "backup");
        Directory.CreateDirectory(backupDirectory);
        string backup = Path.Combine(backupDirectory, "owned.xlsx");
        File.WriteAllText(backup, "owned backup");
        session.TrackFile(backup);
        refused = false;
        try { session.TrackFile(caller); } catch (IOException) { refused = true; }
        Check(refused, "cleanup cannot register an external file");
        refused = false;
        try { session.Output("..\\escape.xlsx"); } catch (ArgumentException) { refused = true; }
        Check(refused, "output cannot traverse outside the probe directory");
        Check(session.Finish(0) == 0 && !File.Exists(copied) && !Directory.Exists(backupDirectory), "success cleans only known generated files and directories");
        Check(File.ReadAllText(concurrent) == "do not delete" && Directory.Exists(concurrentEmptyDirectory), "untracked concurrent file and empty directory survive");
        Check(File.ReadAllText(caller) == "USER CONTENT" && File.Exists(export), "caller input and successful export survive probe cleanup");

        ProbeSession first = Session("same-name", new List<string>(), new List<object>());
        ProbeSession second = Session("same-name", new List<string>(), new List<object>());
        Check(!string.Equals(first.DirectoryPath, second.DirectoryPath, StringComparison.OrdinalIgnoreCase), "concurrent runs use unique directories");
        first.Finish(0); second.Finish(0);
    }

    private static void RepeatedReferences()
    {
        List<object> released = new List<object>();
        ProbeSession session = Session("repeated-rcw", new List<string>(), released);
        object same = new object();
        session.Take(same);
        session.Take(same); // Two COM property acquisitions may return the same RCW.
        Check(session.Finish(0) == 0 && released.Count == 2, "release per acquisition without FinalRelease or deduplication");
    }

    [STAThread]
    private static int Main()
    {
        root = Path.Combine(Path.GetTempPath(), "probe-safety-fixture-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            BorrowedAndSingleInstance();
            PropertySnapshots();
            FailureAndEvidence();
            AttachOnlyAndExitCodes();
            FileIsolation();
            RepeatedReferences();
            Console.WriteLine("LiveProbeSafetyFixture: " + checks + " PASS, 0 FAIL");
            // This fixture owns its entire unique sandbox, including sentinels.
            // Live probes use the stricter registered-file-only cleanup tested above.
            Directory.Delete(root, true);
            return 0;
        }
        catch (Exception error)
        {
            Console.WriteLine("LiveProbeSafetyFixture FAIL: " + error + "\nEvidence: " + root);
            return 1;
        }
    }
}
