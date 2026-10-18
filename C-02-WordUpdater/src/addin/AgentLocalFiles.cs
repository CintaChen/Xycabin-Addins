using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    // Resources are exported under LocalAppData, never beside a shadow-cached DLL.
    internal static class AgentLocalFiles
    {
        internal static string CacheRoot = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        internal static Action<string> BrowserLauncher = LaunchBrowser;
        internal static Func<IntPtr, string, int> OpenWithDialog = ShowOpenWith;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct OpenAsInfo
        {
            [MarshalAs(UnmanagedType.LPWStr)] internal string File;
            [MarshalAs(UnmanagedType.LPWStr)] internal string Class;
            internal uint Flags;
        }
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        private static extern int SHOpenWithDialog(IntPtr parent, ref OpenAsInfo info);
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        private static extern int AssocQueryStringW(uint flags, uint type, string association, string extra, StringBuilder output, ref uint size);

        internal static string Export(string folder, string name, byte[] content)
        {
            string digest;
            using (var sha = SHA256.Create()) digest = BitConverter.ToString(sha.ComputeHash(content)).Replace("-", "").ToLowerInvariant();
            string directory = Path.Combine(CacheRoot, folder, digest);
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, name);
            if (File.Exists(path))
            {
                byte[] existing = File.ReadAllBytes(path);
                bool equal = existing.Length == content.Length;
                for (int i = 0; equal && i < existing.Length; i++) equal = existing[i] == content[i];
                if (equal) return path;
            }
            // A short sibling name avoids exceeding .NET Framework MAX_PATH on deep project paths.
            string temporary = Path.Combine(directory, Guid.NewGuid().ToString("N").Substring(0, 16) + ".tmp");
            try
            {
                File.WriteAllBytes(temporary, content);
                // Another host can export the same shared page concurrently.
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else
                {
                    try { File.Move(temporary, path); }
                    catch (IOException) { if (!File.Exists(path)) throw; File.Replace(temporary, path, null); }
                }
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
            return path;
        }
        internal static string ExportHome()
        {
            using (Stream source = typeof(Connect).Assembly.GetManifestResourceStream("WpsShared.Agent.Home.html"))
            {
                if (source == null) throw new FileNotFoundException("缺少内嵌网站 WpsShared.Agent.Home.html");
                using (var buffer = new MemoryStream())
                {
                    source.CopyTo(buffer);
                    return Export(Path.Combine("Xy", "AgentEvo", "Home"), "index.html", buffer.ToArray());
                }
            }
        }
        internal static string ExportDocument(bool guide, string source)
        {
            return ExportDocumentBundle(typeof(Connect).Assembly.GetName().Name, guide, source,
                ReadDocumentResource("WpsShared.Agent.Assets.xml"), ReadDocumentResource);
        }
        private static byte[] ReadDocumentResource(string name)
        {
            using (Stream stream = typeof(Connect).Assembly.GetManifestResourceStream(name))
            {
                if (stream == null) throw new FileNotFoundException("缺少内嵌文档资源 " + name);
                using (var buffer = new MemoryStream()) { stream.CopyTo(buffer); return buffer.ToArray(); }
            }
        }
        private sealed class DocumentEntry
        {
            internal string Path;
            internal byte[] Bytes;
            internal DocumentEntry(string path, byte[] bytes) { Path = path; Bytes = bytes; }
        }
        private static string Sha256(byte[] bytes)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
        }
        private static string CheckedHash(string value)
        {
            if (value == null || value.Length != 64) throw new InvalidDataException("资源摘要必须为64位十六进制。");
            foreach (char c in value) if (!Uri.IsHexDigit(c)) throw new InvalidDataException("资源摘要不合法。");
            return value.ToLowerInvariant();
        }
        private static void Attributes(XmlElement element, params string[] allowed)
        {
            foreach (XmlAttribute attribute in element.Attributes)
                if (Array.IndexOf(allowed, attribute.Name) < 0) throw new InvalidDataException("资源索引含未知属性 " + attribute.Name);
        }
        private static XmlDocument ReadXml(byte[] bytes)
        {
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                MaxCharactersInDocument = 16 * 1024 * 1024 };
            var document = new XmlDocument { XmlResolver = null };
            using (var memory = new MemoryStream(bytes, false))
            using (XmlReader reader = XmlReader.Create(memory, settings)) document.Load(reader);
            return document;
        }
        private static string CheckedPath(string path, bool document)
        {
            if (string.IsNullOrEmpty(path) || path.IndexOf('\\') >= 0 || path.IndexOf(':') >= 0 || Path.IsPathRooted(path))
                throw new InvalidDataException("资源路径必须是包内正斜杠相对路径。");
            foreach (string part in path.Split('/'))
                if (part.Length == 0 || part == "." || part == ".." || part.TrimEnd(' ', '.') != part || part.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                    throw new InvalidDataException("资源路径包含非法目录或文件名。");
            if (document)
            {
                if (path != "AP3-说明/开发日志.md" && path != "AP3-说明/使用说明.md") throw new InvalidDataException("文档路径不在允许位置。");
            }
            else if (!(path.StartsWith("AP2-美工/", StringComparison.Ordinal) && path.EndsWith(".png", StringComparison.Ordinal))
                && !(path.StartsWith("src/documentation/", StringComparison.Ordinal) && path.EndsWith(".svg", StringComparison.Ordinal)))
                throw new InvalidDataException("仅允许AP2 PNG与documentation SVG资源。");
            return path;
        }
        private static byte[] ResourceBytes(Func<string, byte[]> read, string resource, string hash)
        {
            byte[] bytes = read(resource);
            if (bytes == null || Sha256(bytes) != CheckedHash(hash)) throw new InvalidDataException("资源缺失或摘要不符 " + resource);
            return bytes;
        }
        private static void ValidateImage(string path, byte[] bytes)
        {
            if (path.EndsWith(".png", StringComparison.Ordinal))
            {
                byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
                if (bytes.Length < signature.Length) throw new InvalidDataException("PNG资源不完整。");
                for (int i = 0; i < signature.Length; i++) if (bytes[i] != signature[i]) throw new InvalidDataException("资源不是PNG。");
                using (var memory = new MemoryStream(bytes, false))
                using (Image image = Image.FromStream(memory, true, true))
                    if (image.Width <= 0 || image.Height <= 0) throw new InvalidDataException("PNG尺寸不合法。");
            }
            else
            {
                XmlDocument svg = ReadXml(bytes);
                if (svg.DocumentElement == null || svg.DocumentElement.LocalName != "svg" || svg.DocumentElement.NamespaceURI != "http://www.w3.org/2000/svg")
                    throw new InvalidDataException("SVG根元素不合法。");
            }
        }
        private static void CheckInlineAssets(string source, HashSet<string> paths)
        {
            // The build collector supports broader Markdown syntax. Here also
            // reject a missing asset for the local inline links used by our docs.
            var body = new StringBuilder(); char fence = '\0'; int length = 0; bool previousBlank = true, indented = false;
            foreach (string line in Regex.Replace(source, "<!--.*?-->", "", RegexOptions.Singleline).Replace("\r\n", "\n").Split('\n'))
            {
                Match opening = Regex.Match(line, "^ {0,3}([`~]{3,})(.*)$");
                if (opening.Success)
                {
                    string mark = opening.Groups[1].Value;
                    bool uniform = mark.Trim(mark[0]).Length == 0;
                    if (uniform && fence == '\0') { fence = mark[0]; length = mark.Length; continue; }
                    if (uniform && fence == mark[0] && mark.Length >= length && opening.Groups[2].Value.Trim().Length == 0) { fence = '\0'; continue; }
                }
                if (fence == '\0')
                {
                    bool blank = string.IsNullOrWhiteSpace(line), indentation = line.StartsWith("    ", StringComparison.Ordinal) || line.StartsWith("\t", StringComparison.Ordinal);
                    if ((indentation && (previousBlank || indented)) || (blank && indented)) { indented = true; previousBlank = blank; continue; }
                    indented = false; previousBlank = blank; body.AppendLine(line);
                }
            }
            string markdown = Regex.Replace(body.ToString(), "(`+)(.*?)\\1", "", RegexOptions.Singleline);
            foreach (Match link in Regex.Matches(markdown, "(?<!\\\\)!?\\[[^]\\r\\n]*\\]\\(\\s*(<[^>\\r\\n]+>|[^\\s)]+)(?:\\s+[\\\"'][^\\r\\n]*[\\\"'])?\\s*\\)"))
            {
                string target = link.Groups[1].Value.Trim('<', '>'); Uri uri;
                if (Uri.TryCreate(target, UriKind.Absolute, out uri) && (uri.Scheme == "http" || uri.Scheme == "https")) continue;
                target = Uri.UnescapeDataString(target);
                if (!target.EndsWith(".png", StringComparison.OrdinalIgnoreCase) && !target.EndsWith(".svg", StringComparison.OrdinalIgnoreCase)) continue;
                string relative = target.StartsWith("../", StringComparison.Ordinal) ? target.Substring(3) : "AP3-说明/" + target;
                relative = CheckedPath(relative, false);
                if (!paths.Contains(relative)) throw new InvalidDataException("文档引用缺少对应资产 " + relative);
            }
        }
        private static List<DocumentEntry> DocumentEntries(bool guide, string source, byte[] manifest, Func<string, byte[]> read)
        {
            if (source == null || read == null) throw new ArgumentNullException("document source/resource reader");
            XmlDocument index = ReadXml(manifest);
            XmlElement root = index.DocumentElement;
            if (root == null || root.Name != "documentAssets" || root.NamespaceURI.Length != 0 || root.GetAttribute("version") != "1")
                throw new InvalidDataException("文档资源索引版本不支持。");
            Attributes(root, "version");
            string selected = guide ? "UserGuide" : "Readme";
            var names = new HashSet<string>(StringComparer.Ordinal);
            XmlElement chosen = null;
            foreach (XmlNode node in root.ChildNodes)
            {
                XmlElement entry = node as XmlElement;
                if (entry == null) { if (node.NodeType != XmlNodeType.Comment && node.NodeType != XmlNodeType.Whitespace) throw new InvalidDataException("资源索引含非文档内容。"); continue; }
                Attributes(entry, "name", "resource", "path", "sha256");
                string name = entry.GetAttribute("name");
                if (entry.Name != "document" || entry.NamespaceURI.Length != 0 || (name != "Readme" && name != "UserGuide") || !names.Add(name))
                    throw new InvalidDataException("资源索引文档名称无效或重复。");
                string expectedPath = name == "Readme" ? "AP3-说明/开发日志.md" : "AP3-说明/使用说明.md";
                if (CheckedPath(entry.GetAttribute("path"), true) != expectedPath || entry.GetAttribute("resource") != "WpsShared.Agent." + name + ".md")
                    throw new InvalidDataException("文档资源身份不符。");
                CheckedHash(entry.GetAttribute("sha256"));
                var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (XmlNode child in entry.ChildNodes)
                {
                    XmlElement asset = child as XmlElement;
                    if (asset == null) { if (child.NodeType != XmlNodeType.Comment && child.NodeType != XmlNodeType.Whitespace) throw new InvalidDataException("资源索引含非资产内容。"); continue; }
                    Attributes(asset, "path", "resource", "sha256", "size");
                    string path = CheckedPath(asset.GetAttribute("path"), false), hash = CheckedHash(asset.GetAttribute("sha256"));
                    long size;
                    if (asset.Name != "asset" || asset.NamespaceURI.Length != 0 || asset.HasChildNodes || !paths.Add(path)
                        || !long.TryParse(asset.GetAttribute("size"), NumberStyles.None, CultureInfo.InvariantCulture, out size) || size <= 0
                        || asset.GetAttribute("resource") != "WpsShared.Agent.Asset." + hash + Path.GetExtension(path))
                        throw new InvalidDataException("资产索引字段无效、路径重复或资源身份不符。");
                }
                if (name == selected) chosen = entry;
            }
            if (names.Count != 2 || chosen == null) throw new InvalidDataException("资源索引必须包含两类文档。");
            string documentPath = chosen.GetAttribute("path");
            byte[] documentBytes = ResourceBytes(read, chosen.GetAttribute("resource"), chosen.GetAttribute("sha256"));
            byte[] supplied = new UTF8Encoding(false).GetBytes(source);
            if (Sha256(documentBytes) != Sha256(supplied)) throw new InvalidDataException("传入文档与内嵌资源不一致。");
            var result = new List<DocumentEntry> { new DocumentEntry(documentPath, documentBytes) };
            var selectedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (XmlElement asset in chosen.SelectNodes("asset"))
            {
                string path = asset.GetAttribute("path");
                byte[] bytes = ResourceBytes(read, asset.GetAttribute("resource"), asset.GetAttribute("sha256"));
                if (bytes.LongLength != long.Parse(asset.GetAttribute("size"), CultureInfo.InvariantCulture)) throw new InvalidDataException("资产大小不符 " + path);
                ValidateImage(path, bytes);
                selectedPaths.Add(path);
                result.Add(new DocumentEntry(path, bytes));
            }
            CheckInlineAssets(source, selectedPaths);
            result.Sort(delegate(DocumentEntry left, DocumentEntry right) { return string.CompareOrdinal(left.Path, right.Path); });
            return result;
        }
        private static string BundleHash(List<DocumentEntry> entries)
        {
            using (var memory = new MemoryStream())
            {
                using (var writer = new BinaryWriter(memory, Encoding.UTF8, true))
                    foreach (DocumentEntry entry in entries) { writer.Write(entry.Path); writer.Write(entry.Bytes.Length); writer.Write(entry.Bytes); }
                return Sha256(memory.ToArray());
            }
        }
        private static void PlainPath(string root, string path)
        {
            string full = Path.GetFullPath(path), boundary = Path.GetFullPath(root);
            if (!string.Equals(full, boundary, StringComparison.OrdinalIgnoreCase) && !full.StartsWith(boundary + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                throw new IOException("文档缓存目标越界。");
            string current = full;
            while (true)
            {
                if ((File.Exists(current) || Directory.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("文档缓存不允许联接或符号链接。");
                if (string.Equals(current, boundary, StringComparison.OrdinalIgnoreCase)) break;
                current = Path.GetDirectoryName(current);
                if (current == null) throw new IOException("文档缓存路径不完整。");
            }
        }
        private static void AtomicDocumentFile(string root, string path, byte[] bytes)
        {
            PlainPath(root, path);
            string directory = Path.GetDirectoryName(path); Directory.CreateDirectory(directory);
            if (File.Exists(path) && Sha256(File.ReadAllBytes(path)) == Sha256(bytes)) return;
            string temporary = Path.Combine(directory, Guid.NewGuid().ToString("N").Substring(0, 16) + ".tmp");
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None)) stream.Write(bytes, 0, bytes.Length);
                PlainPath(root, path);
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
        // Pure resource-reader seam: no source checkout, window, launcher or COM is needed.
        internal static string ExportDocumentBundle(string assemblyName, bool guide, string source, byte[] manifest, Func<string, byte[]> read)
        {
            if (string.IsNullOrEmpty(assemblyName) || assemblyName.TrimEnd(' ', '.') != assemblyName || assemblyName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                throw new InvalidDataException("文档缓存程序集名称不合法。");
            List<DocumentEntry> entries = DocumentEntries(guide, source, manifest, read);
            string digest = BundleHash(entries);
            string root = Path.GetFullPath(CacheRoot);
            string directory = Path.Combine(root, assemblyName, "AgentDocuments", digest.Substring(0, 32));
            PlainPath(root, directory); Directory.CreateDirectory(directory);
            string lockPath = Path.Combine(directory, ".bundle.lock"); PlainPath(root, lockPath);
            FileStream lease = null;
            for (int attempt = 0; lease == null; attempt++)
            {
                try { lease = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
                catch (IOException) { if (attempt >= 200) throw; Thread.Sleep(15); }
            }
            using (lease)
            {
                string marker = Path.Combine(directory, ".bundle-sha256"); PlainPath(root, marker);
                if (File.Exists(marker))
                {
                    string previous = File.ReadAllText(marker, Encoding.UTF8).Trim();
                    bool hash = previous.Length == 64;
                    foreach (char c in previous) hash = hash && Uri.IsHexDigit(c);
                    if (hash && !string.Equals(previous, digest, StringComparison.OrdinalIgnoreCase))
                        throw new IOException("文档资源包摘要前缀冲突，拒绝复用旧包。");
                }
                foreach (DocumentEntry entry in entries)
                    AtomicDocumentFile(root, Path.Combine(directory, entry.Path.Replace('/', Path.DirectorySeparatorChar)), entry.Bytes);
                // The full digest commits only after all selected document/assets exist.
                AtomicDocumentFile(root, marker, Encoding.ASCII.GetBytes(digest));
            }
            return Path.Combine(directory, "AP3-说明", guide ? "使用说明.md" : "开发日志.md");
        }
        internal static void OpenDocument(IntPtr owner, string path)
        {
            int result = OpenWithDialog(owner, path);
            // S_FALSE and ERROR_CANCELLED are normal cancellation outcomes.
            if (result >= 0 || result == unchecked((int)0x800704C7)) return;
            Marshal.ThrowExceptionForHR(result);
        }
        private static int ShowOpenWith(IntPtr owner, string path)
        {
            var info = new OpenAsInfo { File = path, Class = null, Flags = 0x00000004 | 0x00000020 };
            return SHOpenWithDialog(owner, ref info);
        }
        internal static string DefaultBrowserExecutable()
        {
            // Query the URL protocol, rather than the potentially unrelated .html association.
            foreach (string protocol in new string[] { "https", "http" })
            {
                uint size = 0;
                AssocQueryStringW(0x00001000, 2, protocol, "open", null, ref size);
                if (size == 0 || size > 32768) continue;
                var output = new StringBuilder((int)size);
                if (AssocQueryStringW(0x00001000, 2, protocol, "open", output, ref size) == 0 && File.Exists(output.ToString()))
                    return output.ToString();
            }
            throw new InvalidOperationException("无法查询系统默认浏览器，请在 Windows 设置中选择默认浏览器。");
        }
        private static void LaunchBrowser(string path)
        {
            string url = new Uri(Path.GetFullPath(path)).AbsoluteUri;
            Process.Start(new ProcessStartInfo(DefaultBrowserExecutable(), "\"" + url + "\"") { UseShellExecute = true });
        }
    }
}
