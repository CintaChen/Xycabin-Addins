// ValidateManifest —— 用 GAC 中 MSBuild v4.0 的
// Microsoft.Build.Tasks.Deployment.ManifestUtilities.ManifestReader
// 解析 ClickOnce/VSTO 清单，打印解析模型（标识、依赖、入口点），
// 用于与 CPAHelper 的清单做结构对比。
using System;
using System.IO;
using System.Reflection;

internal static class ValidateManifest
{
    private static readonly string TasksDll = Path.Combine(
        Environment.GetEnvironmentVariable("SystemRoot") ?? @"C:\Windows",
        @"Microsoft.NET\assembly\GAC_MSIL\Microsoft.Build.Tasks.v4.0\v4.0_4.0.0.0__b03f5f7f11d50a3a\Microsoft.Build.Tasks.v4.0.dll");

    private static int Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("usage: ValidateManifest <manifest路径> [...]");
            return 2;
        }
        Assembly asm = Assembly.LoadFrom(TasksDll);
        Type reader = asm.GetType("Microsoft.Build.Tasks.Deployment.ManifestUtilities.ManifestReader");
        if (reader == null) { Console.WriteLine("ManifestReader not found"); return 1; }
        MethodInfo readManifest = reader.GetMethod("ReadManifest",
            BindingFlags.Public | BindingFlags.Static, null,
            new[] { typeof(string), typeof(bool) }, null);
        if (readManifest == null) { Console.WriteLine("ReadManifest(string,bool) not found"); return 1; }

        foreach (string path in args)
        {
            Console.WriteLine("===== " + Path.GetFileName(path) + " =====");
            try
            {
                object manifest = readManifest.Invoke(null, new object[] { path, false });
                DumpObject(manifest, 0, 0);
            }
            catch (TargetInvocationException tie)
            {
                Exception ex = tie.InnerException ?? tie;
                Console.WriteLine("PARSE FAILED: " + ex.GetType().Name + ": " + ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAILED: " + ex.GetType().Name + ": " + ex.Message);
            }
        }
        return 0;
    }

    // 打印关键属性：Identity / EntryPoint / AssemblyReferences / FileReferences
    private static void DumpObject(object obj, int depth, int budget)
    {
        if (obj == null || depth > 3 || budget > 40) return;
        Type t = obj.GetType();
        foreach (string propName in new[] { "Identity", "EntryPointIdentity", "MainFile",
                                            "AssemblyReferences", "FileReferences", "DeployManifestPath" })
        {
            PropertyInfo pi = t.GetProperty(propName,
                BindingFlags.Public | BindingFlags.Instance);
            if (pi == null) continue;
            object val;
            try { val = pi.GetValue(obj, null); }
            catch { continue; }
            if (val == null) { Console.WriteLine(propName + ": (null)"); continue; }
            if (val is System.Collections.IEnumerable && !(val is string))
            {
                int i = 0;
                foreach (object item in (System.Collections.IEnumerable)val)
                {
                    Console.WriteLine(propName + "[" + i + "]: " + Brief(item));
                    i++;
                    if (i >= 30) break;
                }
            }
            else
            {
                Console.WriteLine(propName + ": " + val);
            }
        }
    }

    private static string Brief(object item)
    {
        if (item == null) return "(null)";
        Type t = item.GetType();
        string name = null, codebase = null, size = null;
        try { name = Convert.ToString(t.GetProperty("Identity") != null
            ? t.GetProperty("Identity").GetValue(item, null)
            : (t.GetProperty("Name") != null ? t.GetProperty("Name").GetValue(item, null) : null)); }
        catch { }
        try { if (t.GetProperty("Codebase") != null) codebase = Convert.ToString(t.GetProperty("Codebase").GetValue(item, null)); } catch { }
        try { if (t.GetProperty("Size") != null) size = Convert.ToString(t.GetProperty("Size").GetValue(item, null)); } catch { }
        return string.Format("[{0}] name={1} codebase={2} size={3}", t.Name, name ?? "?", codebase ?? "-", size ?? "-");
    }
}
