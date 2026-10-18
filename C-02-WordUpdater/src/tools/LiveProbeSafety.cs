using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

// Shared only by diagnostic tools, never compiled into an installed add-in.
// CreateInstance can return an existing single-instance WPS Application. A probe
// owns documents it creates, not the Application or any process discovered later.
internal sealed class ProbeSession : IDisposable
{
    private readonly Action<string> log;
    private readonly Func<object, bool> isCom;
    private readonly Action<object> release;
    private readonly List<Action> closes = new List<Action>();
    private readonly List<Action> restores = new List<Action>();
    private readonly List<object> references = new List<object>();
    private readonly List<string> files = new List<string>();
    private readonly List<string> directories = new List<string>();
    private bool disposed, safetyFailed;
    internal readonly string DirectoryPath;
    internal int ExitCode { get; private set; }

    internal ProbeSession(string name) : this(name, Path.GetTempPath(), Console.WriteLine,
        Marshal.IsComObject, delegate(object value) { Marshal.ReleaseComObject(value); }) { }

    // Filesystem and host operations are injectable so the same production
    // lifecycle can be checked without starting or connecting to WPS.
    internal ProbeSession(string name, string parent, Action<string> output,
        Func<object, bool> checkCom, Action<object> releaseReference)
    {
        if (string.IsNullOrEmpty(name) || name != Path.GetFileName(name)) throw new ArgumentException("Invalid probe name.");
        log = output ?? delegate { };
        isCom = checkCom;
        release = releaseReference;
        ExitCode = 1;
        DirectoryPath = Path.Combine(Path.GetFullPath(parent), name + "-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(DirectoryPath);
        directories.Add(DirectoryPath);
        log("PROBE_DIRECTORY: " + DirectoryPath);
    }

    internal object Take(object value)
    {
        if (value != null) references.Add(value); // Each acquisition owns one reference.
        return value;
    }

    internal object Own(object value, Action close)
    {
        if (value == null) throw new InvalidOperationException("Host did not create the probe document.");
        Take(value);
        closes.Add(close);
        return value;
    }

    internal void OnClose(Action close) { closes.Add(close); }

    internal bool Override(Func<object> read, Action<object> write, object temporary, string label)
    {
        object original;
        try { original = read(); }
        catch (Exception error) { log("UNCHANGED " + label + ": " + error.Message); return false; }
        if (original == null || original == DBNull.Value || original == Type.Missing)
        { log("UNCHANGED " + label + ": original value is unavailable"); return false; }
        // Register first: a setter can mutate the host and then throw.
        restores.Add(delegate { write(original); });
        try { write(temporary); return true; }
        catch (Exception error) { log("OVERRIDE_UNAVAILABLE " + label + ": " + error.Message); return false; }
    }

    internal string Output(string filename)
    {
        if (string.IsNullOrEmpty(filename) || filename != Path.GetFileName(filename))
            throw new ArgumentException("Probe output must be a filename inside its own directory.");
        string path = Path.Combine(DirectoryPath, filename);
        if (File.Exists(path) || Directory.Exists(path)) throw new IOException("Output already exists: " + path);
        files.Add(path);
        return path;
    }

    internal string CopyInput(string source, string filename)
    {
        string target = Output(filename);
        File.Copy(Path.GetFullPath(source), target, false);
        return target;
    }

    internal static void RejectExistingOutput(string path)
    {
        if (File.Exists(path) || Directory.Exists(path))
            throw new IOException("Refusing to overwrite an existing output: " + path);
    }

    // Never write test data to an arbitrary caller path. An optional output is
    // exported from the isolated source only after a successful run, without overwrite.
    internal void Export(string source, string target)
    {
        RejectExistingOutput(target);
        File.Copy(source, Path.GetFullPath(target), false);
    }

    internal void TrackFile(string path)
    {
        string full = Path.GetFullPath(path);
        if (!full.StartsWith(DirectoryPath + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            throw new IOException("Generated file is outside the probe directory: " + path);
        if (!files.Contains(full)) files.Add(full);
        // The known generated file may have created subdirectories (e.g. backup).
        // Never enumerate/delete unrelated directories created concurrently.
        string parent = Path.GetDirectoryName(full);
        while (!string.Equals(parent, DirectoryPath, StringComparison.OrdinalIgnoreCase))
        {
            if (!directories.Contains(parent)) directories.Add(parent);
            parent = Path.GetDirectoryName(parent);
        }
    }

    internal void Fail(string message) { safetyFailed = true; ExitCode = 1; log("FAIL: " + message); }

    internal int Finish(int result)
    {
        if (disposed) return ExitCode;
        ExitCode = safetyFailed || (result != 0 && result != 2) ? 1 : result;
        Dispose();
        return ExitCode;
    }

    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        for (int i = closes.Count - 1; i >= 0; i--)
            try { closes[i](); } catch (Exception error) { Fail("closing probe document: " + error.Message); }
        for (int i = restores.Count - 1; i >= 0; i--)
            try { restores[i](); } catch (Exception error) { Fail("restoring host property: " + error.Message); }
        for (int i = references.Count - 1; i >= 0; i--)
            Release(references[i]);
        references.Clear();
        if (ExitCode == 0)
        {
            // Non-recursive removal: untracked files/directories keep the root.
            foreach (string path in files)
            {
                if (ExitCode != 0) break;
                try { EnsurePlainPath(path); if (File.Exists(path)) File.Delete(path); }
                catch (Exception error) { Fail("cleaning probe file: " + error.Message); }
            }
            directories.Sort(delegate(string left, string right) { return right.Length.CompareTo(left.Length); });
            foreach (string path in directories)
            {
                if (ExitCode != 0) break;
                try
                {
                    EnsurePlainPath(path);
                    if (Directory.Exists(path) && Directory.GetFileSystemEntries(path).Length == 0) Directory.Delete(path);
                }
                catch (Exception error) { Fail("cleaning probe directory: " + error.Message); }
            }
        }
        if (Directory.Exists(DirectoryPath)) log("PROBE_EVIDENCE_RETAINED: " + DirectoryPath);
        log(ExitCode == 0 ? "PROBE_RESULT: PASS" : ExitCode == 2 ? "PROBE_RESULT: SKIP (not PASS)" : "PROBE_RESULT: FAIL");
    }

    private void EnsurePlainPath(string path)
    {
        string current = Path.GetFullPath(path);
        while (current != null)
        {
            if ((File.Exists(current) || Directory.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("Refusing cleanup through a reparse point: " + current);
            if (string.Equals(current, DirectoryPath, StringComparison.OrdinalIgnoreCase)) return;
            current = Path.GetDirectoryName(current);
        }
        throw new IOException("Cleanup path escaped the probe directory: " + path);
    }

    internal void Release(object value)
    {
        if (value == null) return;
        try { if (isCom(value)) release(value); }
        catch (Exception error) { Fail("releasing acquired COM reference: " + error.Message); }
    }

    internal ProbeReferences References() { return new ProbeReferences(this); }

    internal static object Connect(string[] ids, bool attachOnly, Func<string, object> attach,
        Func<string, object> create, Action<string> output)
    {
        foreach (string id in ids)
            try { object value = attach(id); if (value != null) return value; } catch (COMException) { }
        if (!attachOnly)
            foreach (string id in ids)
                try { object value = create(id); if (value != null) return value; } catch (COMException error) { output(error.Message); }
        return null;
    }

    internal static object Connect(string[] ids, bool attachOnly)
    {
        return Connect(ids, attachOnly, Marshal.GetActiveObject, delegate(string id) {
            Type type = Type.GetTypeFromProgID(id);
            return type == null ? null : Activator.CreateInstance(type);
        }, Console.WriteLine);
    }
}

internal sealed class ProbeReferences : IDisposable
{
    private readonly ProbeSession session;
    private readonly List<object> values = new List<object>();
    internal ProbeReferences(ProbeSession owner) { session = owner; }
    internal object Take(object value) { if (value != null) values.Add(value); return value; }
    public void Dispose()
    {
        for (int i = values.Count - 1; i >= 0; i--) session.Release(values[i]);
        values.Clear();
    }
}
