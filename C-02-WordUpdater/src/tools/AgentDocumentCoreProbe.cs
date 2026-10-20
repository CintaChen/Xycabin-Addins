using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

// No forms, shell launcher, browser, COM or WPS. Tests actual export code through
// its resource-reader seam, then checks the real embedded bundle and source bytes.
internal static class AgentDocumentCoreProbe
{
    private const BindingFlags All = BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public;
    private static Assembly assembly;
    private static Type files;
    private static int checks;
    private static string temporary;
    private static void Check(bool value, string label)
    { if (!value) throw new InvalidOperationException(label); checks++; Console.WriteLine("PASS: " + label); }
    private static object Invoke(string name, params object[] args)
    {
        try { return files.GetMethod(name, All).Invoke(null, args); }
        catch (TargetInvocationException error) { throw error.InnerException; }
    }
    private static void Reject(Action action, string label)
    {
        bool rejected = false;
        try { action(); } catch { rejected = true; }
        Check(rejected, label);
    }
    private static string Hash(byte[] bytes)
    { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(); }
    private static byte[] Utf8(string text) { return new UTF8Encoding(false).GetBytes(text); }
    private static byte[] Resource(string name)
    {
        using (Stream stream = assembly.GetManifestResourceStream(name))
        {
            if (stream == null) throw new FileNotFoundException("Missing embedded resource: " + name);
            using (var memory = new MemoryStream()) { stream.CopyTo(memory); return memory.ToArray(); }
        }
    }
    private static XmlDocument Xml(byte[] bytes)
    {
        var doc = new XmlDocument { XmlResolver = null };
        using (var memory = new MemoryStream(bytes, false))
        using (var reader = XmlReader.Create(memory, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) doc.Load(reader);
        return doc;
    }
    private static string Bundle(string markdown)
    { return Directory.GetParent(Path.GetDirectoryName(markdown)).FullName; }
    private static string Target(string bundle, string relative)
    { return Path.Combine(bundle, relative.Replace('/', Path.DirectorySeparatorChar)); }
    private sealed class Sample
    {
        internal const string Png = "C-01-ExcelIterate/AB-Graphic/图像 样例.png", Svg = "C-01-ExcelIterate/AB-Graphic/图例 测试.svg";
        internal string Developer = "# 开发\r\n![插图](<../C-01-ExcelIterate/AB-Graphic/图像 样例.png>)\r\n[矢量](<../C-01-ExcelIterate/AB-Graphic/图例 测试.svg>)\r\n";
        internal readonly string Guide = "# 用户\r\n仅用户说明\r\n";
        internal readonly Dictionary<string, byte[]> Resources = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        internal readonly XmlDocument Manifest = new XmlDocument();
        internal Sample(int color)
        {
            byte[] png;
            using (var bitmap = new Bitmap(3, 2))
            using (var memory = new MemoryStream()) { bitmap.SetPixel(0, 0, Color.FromArgb(color, 42, 90)); bitmap.Save(memory, ImageFormat.Png); png = memory.ToArray(); }
            Manifest.AppendChild(Manifest.CreateElement("documentAssets")); Manifest.DocumentElement.SetAttribute("version", "1");
            XmlElement developer = Document("Readme", Developer, "C-04-Devlog/B-Devlog.md");
            Asset(developer, Png, png); Asset(developer, Svg, Utf8("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"3\" height=\"2\"><path d=\"M0 0L3 2\"/></svg>"));
            Document("UserGuide", Guide, "C-04-Devlog/A-UserGuide.md");
        }
        private XmlElement Document(string name, string text, string path)
        {
            string resource = "WpsShared.Agent." + name + ".md"; byte[] bytes = Utf8(text); Resources.Add(resource, bytes);
            XmlElement entry = Manifest.CreateElement("document"); entry.SetAttribute("name", name); entry.SetAttribute("resource", resource);
            entry.SetAttribute("path", path); entry.SetAttribute("sha256", Hash(bytes)); Manifest.DocumentElement.AppendChild(entry); return entry;
        }
        private void Asset(XmlElement owner, string path, byte[] bytes)
        {
            XmlElement entry = Manifest.CreateElement("asset"); owner.AppendChild(entry); Set(entry, path, bytes);
        }
        internal void Set(XmlElement entry, string path, byte[] bytes)
        {
            string hash = Hash(bytes), resource = "WpsShared.Agent.Asset." + hash + Path.GetExtension(path);
            entry.SetAttribute("path", path); entry.SetAttribute("sha256", hash); entry.SetAttribute("resource", resource);
            entry.SetAttribute("size", bytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)); Resources[resource] = bytes;
        }
        internal XmlElement FirstAsset { get { return (XmlElement)Manifest.SelectSingleNode("/documentAssets/document[@name='Readme']/asset"); } }
        internal byte[] Bytes { get { return Utf8(Manifest.OuterXml); } }
        internal void SetDeveloper(string text)
        {
            Developer = text; Resources["WpsShared.Agent.Readme.md"] = Utf8(text);
            ((XmlElement)Manifest.SelectSingleNode("/documentAssets/document[@name='Readme']")).SetAttribute("sha256", Hash(Utf8(text)));
        }
        internal byte[] Read(string name) { byte[] value; if (!Resources.TryGetValue(name, out value)) throw new FileNotFoundException("injected missing resource"); return value; }
    }
    private static string Export(Sample sample, bool guide)
    {
        return (string)Invoke("ExportDocumentBundle", "CoreFixture", guide, guide ? sample.Guide : sample.Developer, sample.Bytes,
            new Func<string, byte[]>(sample.Read));
    }
    private static void FixtureCases()
    {
        var sample = new Sample(12);
        string oldCurrent = Environment.CurrentDirectory;
        string developer;
        try { Environment.CurrentDirectory = temporary; developer = Export(sample, false); }
        finally { Environment.CurrentDirectory = oldCurrent; }
        string bundle = Bundle(developer);
        Check(File.ReadAllText(developer, Encoding.UTF8) == sample.Developer, "offline export preserves Chinese Markdown and relative URLs without workspace");
        Check(Path.GetFileName(bundle).Length == 32 && File.ReadAllText(Path.Combine(bundle, ".bundle-sha256")).Length == 64, "short cache name retains full digest marker");
        Check(File.Exists(Target(bundle, Sample.Png)) && File.Exists(Target(bundle, Sample.Svg)), "PNG and SVG retain original relative source directories");
        Check(Export(sample, false) == developer, "repeat export reuses the same complete bundle");
        File.WriteAllText(developer, "damaged"); File.WriteAllBytes(Target(bundle, Sample.Png), new byte[] { 0, 1 }); File.Delete(Target(bundle, Sample.Svg));
        Check(Export(sample, false) == developer && File.ReadAllText(developer, Encoding.UTF8) == sample.Developer, "damaged Markdown is repaired without changing bundle identity");
        Check(Hash(File.ReadAllBytes(Target(bundle, Sample.Png))) == sample.FirstAsset.GetAttribute("sha256") && File.Exists(Target(bundle, Sample.Svg)), "damaged PNG and missing SVG are repaired from resource bytes");
        File.WriteAllText(Path.Combine(bundle, "sentinel.txt"), "keep"); Export(sample, false);
        Check(File.ReadAllText(Path.Combine(bundle, "sentinel.txt")) == "keep", "repair does not remove untracked or older cache files");
        Task<string> a = Task.Factory.StartNew(() => Export(sample, false)), b = Task.Factory.StartNew(() => Export(sample, false));
        Task.WaitAll(a, b); Check(a.Result == developer && b.Result == developer, "concurrent export shares a complete atomic bundle");
        Check(Directory.GetFiles(bundle, "*.tmp", SearchOption.AllDirectories).Length == 0, "successful atomic export leaves no sibling temporary files");
        var changed = new Sample(13); string updated = Export(changed, false);
        Check(updated != developer && File.Exists(developer), "asset-only update changes bundle identity and preserves the previous package");
        var reads = new List<string>();
        string guide = (string)Invoke("ExportDocumentBundle", "CoreFixture", true, sample.Guide, sample.Bytes,
            new Func<string, byte[]>(name => { reads.Add(name); return sample.Read(name); }));
        Check(reads.Count == 1 && reads[0] == "WpsShared.Agent.UserGuide.md", "guide export never reads developer Markdown or developer assets");
        Check(!File.Exists(Path.Combine(Bundle(guide), "C-04-Devlog", "B-Devlog.md")), "guide package does not leak developer files or assets");
        var examples = new Sample(12);
        foreach (XmlNode node in examples.Manifest.SelectNodes("//asset")) node.ParentNode.RemoveChild(node);
        examples.SetDeveloper("# 示例\r\n<!--![图](../C-01-ExcelIterate/AB-Graphic/缺失.png)-->\r\n`[代码](../C-01-ExcelIterate/AB-Graphic/缺失.png)`\r\n```md\r\n![围栏](../C-01-ExcelIterate/AB-Graphic/缺失.png)\r\n```\r\n\r\n    [缩进](../C-01-ExcelIterate/AB-Graphic/缺失.png)\r\n");
        Check(File.Exists(Export(examples, false)), "comment, code-span, fence and indented examples do not require image resources");

        string cache = (string)files.GetField("CacheRoot", All).GetValue(null);
        string deep = Path.Combine(temporary, new string('d', Math.Max(1, 132 - temporary.Length - 1)));
        Directory.CreateDirectory(deep); files.GetField("CacheRoot", All).SetValue(null, deep);
        try { string path = Export(sample, false); Check(path.Length < 260 && Target(Bundle(path), Sample.Svg).Length < 260, "132-character cache root remains within Framework path limits"); }
        finally { files.GetField("CacheRoot", All).SetValue(null, cache); }
        int before = Directory.GetFiles(cache, "*", SearchOption.AllDirectories).Length;
        Reject(() => Invoke("ExportDocumentBundle", "CoreFixture", false, sample.Developer, sample.Bytes,
            new Func<string, byte[]>(name => name == sample.FirstAsset.GetAttribute("resource") ? null : sample.Read(name))), "missing embedded PNG is rejected before cache writes");
        Reject(() => Invoke("ExportDocumentBundle", "CoreFixture", false, sample.Developer + "changed", sample.Bytes,
            new Func<string, byte[]>(sample.Read)), "mismatched supplied Markdown is rejected");
        var absent = new Sample(12); absent.FirstAsset.ParentNode.RemoveChild(absent.FirstAsset);
        Reject(() => Export(absent, false), "Markdown PNG reference missing from manifest is rejected");
        var badPng = new Sample(12); badPng.Set(badPng.FirstAsset, Sample.Png, Utf8("not a PNG"));
        Reject(() => Export(badPng, false), "matching hash cannot make undecodable PNG valid");
        var badSvg = new Sample(12); badSvg.Set((XmlElement)badSvg.Manifest.SelectSingleNode("//asset[contains(@path,'.svg')]"), Sample.Svg,
            Utf8("<!DOCTYPE svg [<!ENTITY x SYSTEM 'file:///never-read'>]><svg xmlns=\"http://www.w3.org/2000/svg\">&x;</svg>"));
        Reject(() => Export(badSvg, false), "SVG DTD and external entity are prohibited");
        foreach (string invalid in new[] { "../outside.png", "C:/outside.png", "/C-01-ExcelIterate/AB-Graphic/x.png", "C-01-ExcelIterate/AB-Graphic/../x.png", "C-01-ExcelIterate\\AB-Graphic\\x.png", "C-01-ExcelIterate/AB-Graphic//x.png", "src/x.svg", "C-01-ExcelIterate/AB-Graphic/x.png?query", "AB-Graphic/x.png", "C-09-Unknown/AB-Graphic/x.png" })
        {
            var bad = new Sample(12); bad.FirstAsset.SetAttribute("path", invalid); Reject(() => Export(bad, false), "illegal asset path rejected: " + invalid);
        }
        var duplicate = new Sample(12); duplicate.FirstAsset.ParentNode.AppendChild(duplicate.FirstAsset.CloneNode(true));
        Reject(() => Export(duplicate, false), "duplicate output path is rejected");
        var badSize = new Sample(12); badSize.FirstAsset.SetAttribute("size", "0"); Reject(() => Export(badSize, false), "nonpositive asset size is rejected");
        var badHash = new Sample(12); badHash.FirstAsset.SetAttribute("sha256", new string('a', 64)); Reject(() => Export(badHash, false), "asset resource identity or digest mismatch is rejected");
        Reject(() => Invoke("ExportDocumentBundle", "../outside", false, sample.Developer, sample.Bytes,
            new Func<string, byte[]>(sample.Read)), "assembly cache name cannot escape root");
        Check(Directory.GetFiles(cache, "*", SearchOption.AllDirectories).Length == before, "invalid manifests/resources do not publish partial cache files");
        string marker = Path.Combine(bundle, ".bundle-sha256"), original = File.ReadAllText(marker);
        try { File.WriteAllText(marker, new string('b', 64)); Reject(() => Export(sample, false), "full digest marker detects prefix collision instead of reusing another bundle"); }
        finally { File.WriteAllText(marker, original); }
        File.Delete(marker); Check(Export(sample, false) == developer && File.ReadAllText(marker) == original, "missing digest commit is repaired");
    }
    private static byte[] SourceDocument(string path)
    {
        string text = File.ReadAllText(path, Encoding.UTF8).Replace("\r\n", "\n").Replace("\r", "\n");
        
        return Utf8(text.Replace("\n", Environment.NewLine));
    }
    private static void ActualResources(string project)
    {
        XmlDocument index = Xml(Resource("WpsShared.Agent.Assets.xml"));
        Check(index.DocumentElement.Name == "documentAssets" && index.DocumentElement.GetAttribute("version") == "1", "real embedded assets manifest is version one");
        foreach (XmlElement doc in index.SelectNodes("/documentAssets/document"))
        {
            bool guide = doc.GetAttribute("name") == "UserGuide";
            byte[] source = Resource(doc.GetAttribute("resource"));
            Check(Hash(source) == doc.GetAttribute("sha256"), "real document resource matches manifest digest: " + doc.GetAttribute("name"));
            Check(Hash(SourceDocument(Path.Combine(project, "build", doc.GetAttribute("name") == "Readme" ? "ComposedReadme.md" : "ComposedUserGuide.md"))) == Hash(source), "embedded Markdown matches current normalized source: " + doc.GetAttribute("name"));
            string output = (string)Invoke("ExportDocument", guide, Encoding.UTF8.GetString(source));
            Check(Hash(File.ReadAllBytes(output)) == Hash(source), "actual export preserves original embedded MD bytes: " + doc.GetAttribute("name"));
            foreach (XmlElement asset in doc.SelectNodes("asset"))
            {
                byte[] bytes = Resource(asset.GetAttribute("resource")); string relative = asset.GetAttribute("path");
                Check(Hash(bytes) == asset.GetAttribute("sha256") && bytes.LongLength.ToString(System.Globalization.CultureInfo.InvariantCulture) == asset.GetAttribute("size"), "real asset manifest byte identity: " + relative);
                Check(Hash(File.ReadAllBytes(Target(Path.GetDirectoryName(project), relative))) == Hash(bytes), "embedded asset matches original source: " + relative);
                Check(Hash(File.ReadAllBytes(Target(Bundle(output), relative))) == Hash(bytes), "offline exported asset matches embedded bytes: " + relative);
                if (relative.EndsWith(".png", StringComparison.Ordinal))
                    using (var memory = new MemoryStream(bytes, false)) using (Image image = Image.FromStream(memory, true, true)) Check(image.Width > 0 && image.Height > 0, "real PNG is decodable: " + relative);
                else Check(Xml(bytes).DocumentElement.LocalName == "svg", "real SVG is valid offline XML: " + relative);
            }
            if (guide) Check(!File.Exists(Path.Combine(Bundle(output), "C-04-Devlog", "B-Devlog.md")), "real guide package does not include developer Markdown");
        }
    }
    private static void AccessFlow()
    {
        Type access = assembly.GetType(assembly.GetName().Name + ".AgentDocumentAccess", true);
        MethodInfo verify = access.GetMethod("VerifyPassword", All), open = access.GetMethod("OpenFullText", All);
        Check(!(bool)verify.Invoke(null, new object[] { null }) && !(bool)verify.Invoke(null, new object[] { "known-wrong-fixture-value" }), "password null/wrong are rejected without recording input");
        Check((bool)verify.Invoke(null, new object[] { "Xymircai" }), "confirmed developer value matches the product digest");
        Check((bool)verify.Invoke(null, new object[] { " \tXymircai\r\n" }), "password ignores leading and trailing whitespace");
        Check(!(bool)verify.Invoke(null, new object[] { "xymircai" }) && !(bool)verify.Invoke(null, new object[] { "XYMIRCAI" }), "password comparison remains case sensitive");
        Check(!(bool)verify.Invoke(null, new object[] { "" }) && !(bool)verify.Invoke(null, new object[] { " \t\r\n" }), "empty and whitespace-only password values are rejected");
        int auth = 0, exports = 0, launches = 0;
        Func<bool> cancel = () => { auth++; return false; }; Func<string> export = () => { exports++; return "fixture.md"; }; Action<string> launch = path => { launches++; };
        Check(!(bool)open.Invoke(null, new object[] { false, cancel, export, launch }) && auth == 1 && exports == 0 && launches == 0, "developer cancellation cannot export or launch");
        Func<bool> accept = () => { auth++; return true; };
        Check((bool)open.Invoke(null, new object[] { false, accept, export, launch }) && auth == 2 && exports == 1 && launches == 1, "developer authorization precedes export and launch");
        Check((bool)open.Invoke(null, new object[] { false, accept, export, launch }) && auth == 3 && exports == 2 && launches == 2, "each repeated developer request authorizes again");
        Check((bool)open.Invoke(null, new object[] { true, cancel, export, launch }) && auth == 3 && exports == 3 && launches == 3, "guide bypasses developer authorization only");
        Func<bool> broken = () => { throw new InvalidOperationException("injected authorization"); };
        Check(!(bool)open.Invoke(null, new object[] { false, broken, export, launch }) && exports == 3 && launches == 3, "authorization exception blocks export and launch");
        Func<string> failedExport = () => { throw new IOException("injected export"); };
        Check(!(bool)open.Invoke(null, new object[] { false, accept, failedExport, launch }) && launches == 3, "export exception cannot launch incomplete document");
    }
    private static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        object originalCache = null, originalBrowser = null, originalChooser = null;
        try
        {
            if (args.Length > 2) throw new ArgumentException("Usage: AgentDocumentCoreProbe.exe [addin.dll [project-root]]");
            string dll = args.Length > 0 ? Path.GetFullPath(args[0]) : null;
            if (dll == null)
                foreach (string name in new[] { "WordUpdater.dll", "Excel_Iterate.dll" })
                { string candidate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, name); if (File.Exists(candidate)) { if (dll != null) throw new ArgumentException("Specify the intended add-in DLL."); dll = candidate; } }
            if (dll == null) throw new FileNotFoundException("Specify the add-in DLL.");
            string project = args.Length > 1 ? Path.GetFullPath(args[1]) : Directory.GetParent(Path.GetDirectoryName(dll)).FullName;
            assembly = Assembly.LoadFrom(dll); files = assembly.GetType(assembly.GetName().Name + ".AgentLocalFiles", true);
            originalCache = files.GetField("CacheRoot", All).GetValue(null); originalBrowser = files.GetField("BrowserLauncher", All).GetValue(null); originalChooser = files.GetField("OpenWithDialog", All).GetValue(null);
            temporary = Path.Combine(Path.GetTempPath(), "agent-document-core-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(temporary);
            files.GetField("CacheRoot", All).SetValue(null, Path.Combine(temporary, "缓存 中文"));
            files.GetField("BrowserLauncher", All).SetValue(null, new Action<string>(path => { throw new Exception("Core probe must not launch a browser."); }));
            files.GetField("OpenWithDialog", All).SetValue(null, new Func<IntPtr, string, int>((owner, path) => { throw new Exception("Core probe must not show application chooser."); }));
            FixtureCases(); ActualResources(project); AccessFlow();
            Console.WriteLine("AgentDocumentCoreProbe: " + checks + " PASS, 0 FAIL; resource/cache core only, no forms, launcher, COM or WPS"); return 0;
        }
        catch (Exception error) { Console.WriteLine("FAIL: " + error); return 1; }
        finally
        {
            if (files != null && originalCache != null)
            { files.GetField("CacheRoot", All).SetValue(null, originalCache); files.GetField("BrowserLauncher", All).SetValue(null, originalBrowser); files.GetField("OpenWithDialog", All).SetValue(null, originalChooser); }
            if (temporary != null)
            {
                string full = Path.GetFullPath(temporary), root = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                if (full.StartsWith(root, StringComparison.OrdinalIgnoreCase) && Path.GetFileName(full).StartsWith("agent-document-core-", StringComparison.Ordinal) && Directory.Exists(full)) Directory.Delete(full, true);
            }
        }
    }
}
