using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

public sealed class FakeAgentRibbon
{
    public int Local, Full;
    public bool FailLocal, FailFull;
    public void InvalidateControl(string id) { Local++; if (FailLocal) throw new InvalidOperationException("local refresh failed"); }
    public void Invalidate() { Full++; if (FailFull) throw new InvalidOperationException("full refresh failed"); }
}

internal static class AgentExtensionProbe
{
    private static string CorePage(string text)
    {
        const string pattern="(<script id=\"installer-assets\" type=\"application/json\">)(.*?)(</script>)";
        if(Regex.Matches(text,pattern,RegexOptions.Singleline).Count!=1)throw new Exception("Unique installer data area required");
        return Regex.Replace(text,pattern,m=>m.Groups[1].Value+"{\"packages\":{}}"+m.Groups[3].Value,RegexOptions.Singleline);
    }


    private const BindingFlags All = BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
    private static int passed;
    private static void Check(bool condition, string text) { if (!condition) throw new Exception(text); passed++; Console.WriteLine("PASS: " + text); Console.Out.Flush(); }
    private static object Call(object value, string name, params object[] args) { return value.GetType().GetMethod(name, All).Invoke(value, args); }
    private static bool Visible(object value) { return (bool)value.GetType().GetProperty("Visible", All).GetValue(value, null); }
    private static IEnumerable<Control> Controls(Control parent)
    { foreach (Control c in parent.Controls) { yield return c; foreach (Control child in Controls(c)) yield return child; } }

    private static string FirstBodyLine(string markdown)
    {
        foreach (string line in markdown.Replace("\r\n", "\n").Split('\n'))
        {
            string text = line.Trim();
            if (text.Length == 0 || text.StartsWith("<!--") || text.StartsWith("```") || Regex.IsMatch(text, "^#{1,6}\\s")) continue;
            return Regex.Replace(text, "\\[([^\\]]+)\\]\\(([^)]+)\\)", "$1（$2）").Replace("**", "").Replace("`", "");
        }
        throw new InvalidOperationException("Overview has no body text");
    }

    private static void CheckOverviewFormatting(Form form, RichTextBox box, string overview)
    {
        string body = FirstBodyLine(overview); int start = box.Text.IndexOf(body, StringComparison.Ordinal);
        Check(start >= 0, "actual overview body is rendered");
        box.Select(start, Math.Min(body.Length, 4));
        Check(box.SelectionFont != null && !box.SelectionFont.Bold && box.SelectionColor == Color.FromArgb(36, 50, 71), "actual overview body has regular dark text");
        FieldInfo source = form.GetType().GetField("overview", All);
        try
        {
            // Exercise heading styling independently of the authored overview's structure.
            source.SetValue(form, "# 样例标题\n样例正文"); Call(form, "SetView", false);
            box.Select(box.Text.IndexOf("样例标题", StringComparison.Ordinal), 2);
            Check(box.SelectionFont != null && box.SelectionFont.Bold && box.SelectionColor == Color.FromArgb(37, 99, 235), "explicit Markdown heading has blue bold text");
            box.Select(box.Text.IndexOf("样例正文", StringComparison.Ordinal), 2);
            Check(box.SelectionFont != null && !box.SelectionFont.Bold && box.SelectionColor == Color.FromArgb(36, 50, 71), "explicit Markdown body has regular dark text");
        }
        finally { source.SetValue(form, overview); Call(form, "SetView", false); }
    }

    private static void CheckPasswordForm(Assembly asm)
    {
        string prefix = asm.GetName().Name + ".";
        Type passwordType = asm.GetType(prefix + "DeveloperPasswordForm", true);
        using (Form form = (Form)Activator.CreateInstance(passwordType, true))
        {
            var password = (TextBox)passwordType.GetField("Password", All).GetValue(form);
            var hint = (Label)passwordType.GetField("Hint", All).GetValue(form);
            var error = (Label)passwordType.GetField("Error", All).GetValue(form);
            Check(password.Name == "developerPassword" && password.UseSystemPasswordChar, "developer input is named and masked");
            Check(hint.Name == "passwordHint" && hint.Text == "请输入开发者的微信名称", "developer hint is exact");
            Check(form.AcceptButton != null && ((Control)form.AcceptButton).Text == "确认" && form.CancelButton != null, "confirm and cancel keyboard buttons wired");
            int attempts = 0;
            passwordType.GetField("Verify", All).SetValue(form, new Func<string, bool>(delegate { attempts++; return false; }));
            password.Text = "fixture-invalid"; Call(form, "Submit");
            Check(attempts == 1 && !form.IsDisposed && form.DialogResult != DialogResult.OK && error.Text == "名称不正确，请重试。", "wrong input keeps form available for retry");
            Check(password.Text.Length == 0, "wrong input is cleared, not retained");
            passwordType.GetField("Verify", All).SetValue(form, new Func<string, bool>(delegate { attempts++; throw new InvalidOperationException("fixture-only"); }));
            password.Text = "fixture-exception"; Call(form, "Submit");
            Check(attempts == 2 && !form.IsDisposed && error.Text == "验证未完成，请重试。" && password.Text.Length == 0, "verification exception stays inline without input or exception message");
            passwordType.GetField("Verify", All).SetValue(form, new Func<string, bool>(delegate { attempts++; return true; }));
            password.Text = "fixture-success"; object[] enter = { new Message(), Keys.Enter };
            Check((bool)Call(form, "ProcessCmdKey", enter) && attempts == 3 && form.DialogResult == DialogResult.OK, "Enter submits retry and only successful verification accepts");
            Check(password.Text.Length == 0, "accepted input is cleared on close");
        }
        foreach (string close in new[] { "escape", "cancel", "window" })
        {
            using (Form form = (Form)Activator.CreateInstance(passwordType, true))
            {
                int attempts = 0;
                passwordType.GetField("Verify", All).SetValue(form, new Func<string, bool>(delegate { attempts++; return true; }));
                var password = (TextBox)passwordType.GetField("Password", All).GetValue(form); password.Text = "fixture-cancel";
                if (close == "escape") Call(form, "ProcessCmdKey", new Message(), Keys.Escape);
                else if (close == "cancel") ((Control)form.CancelButton).GetType().GetMethod("OnClick", All).Invoke(form.CancelButton, new object[] { EventArgs.Empty });
                else form.Close();
                Check(attempts == 0 && form.DialogResult != DialogResult.OK && password.Text.Length == 0, close + " cancels and clears input without verification");
            }
        }
        Type catalog = asm.GetType(prefix + "DialogCatalog", true), factory = asm.GetType(prefix + "DialogPreviewFactory", true);
        object definition = catalog.GetMethod("Find", All).Invoke(null, new object[] { "agent.developerpassword" });
        Check((Size)definition.GetType().GetField("Default", All).GetValue(definition) == new Size(420, 230) && (Size)definition.GetType().GetField("Minimum", All).GetValue(definition) == new Size(380, 210), "password window default and minimum registered");
        using (Form preview = (Form)factory.GetMethod("Create", All).Invoke(null, new object[] { "agent.developerpassword", new Size(420, 230) }))
        {
            int attempts = 0;
            passwordType.GetField("Verify", All).SetValue(preview, new Func<string, bool>(delegate { attempts++; return true; }));
            Call(preview, "Submit");
            Check(attempts == 0 && !preview.Visible && preview.DialogResult != DialogResult.OK, "size preview never verifies or opens developer text");
            Check(((TextBox)passwordType.GetField("Password", All).GetValue(preview)).ReadOnly && !((Control)preview.AcceptButton).Enabled, "password size preview disables input and confirmation");
        }
        Type tree = asm.GetType(prefix + "DialogTree", true);
        object roots = tree.GetMethod("Build", All).Invoke(null, null);
        int passwordNodes = 0;
        foreach (object node in (System.Collections.IEnumerable)tree.GetMethod("All", All).Invoke(null, new[] { roots }))
            if ((string)node.GetType().GetField("DialogId", All).GetValue(node) == "agent.developerpassword")
            {
                passwordNodes++; object parent = node.GetType().GetField("Parent", All).GetValue(node);
                Check(parent != null && (string)parent.GetType().GetField("DialogId", All).GetValue(parent) == "agent.developer", "password is a child window of developer preview, not a Ribbon command");
            }
        Check(passwordNodes == 1, "password dialog appears exactly once in the window tree");
    }

    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        string temporary = null;
        try
        {
            Assembly asm = Assembly.LoadFrom(Path.GetFullPath(args[0])); string prefix = asm.GetName().Name + ".";
            Type state = asm.GetType(prefix + "AgentExtensionController", true);
            var logs = new List<string>(); bool saved = false; int saves = 0;
            Action<string> log = delegate(string text) { logs.Add(text); };
            object controller = Activator.CreateInstance(state, All, null, new object[] { new Func<bool>(delegate { return false; }), new Action<bool>(delegate(bool v) { saved = v; saves++; }), log }, null);
            Check(!Visible(controller), "first launch hidden");
            var first = new FakeAgentRibbon(); var second = new FakeAgentRibbon();
            Call(controller, "Register", first); Call(controller, "Register", first); Call(controller, "Register", second);
            Check((int)state.GetProperty("RibbonCount", All).GetValue(controller, null) == 2, "duplicate registration");
            Check((string)Call(controller, "SetVisible", true) == "" && Visible(controller) && saved && saves == 1, "show and persist");
            Check(first.Local == 2 && second.Local == 2, "all registered ribbons refreshed");
            Check((string)Call(controller, "SetVisible", false) == "" && !Visible(controller) && !saved, "hide and persist");
            first.FailLocal = true;
            Check((string)Call(controller, "SetVisible", true) == "" && first.Full == 1, "full refresh fallback");
            second.FailLocal = second.FailFull = true;
            int before = saves;
            Check(((string)Call(controller, "SetVisible", false)).Contains("恢复原") && Visible(controller) && saves == before, "refresh failure rolls back without saving");
            Check(first.Full >= 3 && logs.Count > 0, "rollback refresh and diagnostics");
            Call(controller, "Clear");
            Check((int)state.GetProperty("RibbonCount", All).GetValue(controller, null) == 0, "shutdown drops ribbon references");
            Check(((string)Call(controller, "SetVisible", false)).Contains("尚未取得") && Visible(controller), "unloaded ribbon cannot report success");
            object corrupt = Activator.CreateInstance(state, All, null, new object[] { new Func<bool>(delegate { throw new IOException("corrupt"); }), new Action<bool>(delegate(bool v) { throw new IOException("cannot save"); }), log }, null);
            Check(!Visible(corrupt), "corrupt settings defaults hidden"); Call(corrupt, "Register", new FakeAgentRibbon());
            Check(((string)Call(corrupt, "SetVisible", true)).Contains("未能记住") && Visible(corrupt), "save failure retains current state");
            object restored = Activator.CreateInstance(state, All, null, new object[] { new Func<bool>(delegate { return true; }), new Action<bool>(delegate(bool v) { }), log }, null);
            Check(Visible(restored), "next launch restoration");
            Type settings = asm.GetType(prefix + "AgentSettings", true);
            temporary = Path.Combine(Path.GetTempPath(), "agent-probe-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(temporary);
            string exportRoot = Path.Combine(temporary, new string('p', Math.Max(1, 132 - temporary.Length - 1)));
            Type files = asm.GetType(prefix + "AgentLocalFiles", true);
            files.GetField("CacheRoot", All).SetValue(null, exportRoot);
            string path = Path.Combine(temporary, "AgentExtension.state");
            Check(!(bool)settings.GetMethod("Read", All).Invoke(null, new object[] { path }), "missing state");
            settings.GetMethod("Save", All).Invoke(null, new object[] { path, true });
            Check((bool)settings.GetMethod("Read", All).Invoke(null, new object[] { path }), "UTF8 state round trip");
            settings.GetMethod("Save", All).Invoke(null, new object[] { path, false });
            Check(!(bool)settings.GetMethod("Read", All).Invoke(null, new object[] { path }), "atomic replacement");
            File.WriteAllText(path, "damaged"); bool rejected = false;
            try { settings.GetMethod("Read", All).Invoke(null, new object[] { path }); } catch (TargetInvocationException) { rejected = true; }
            Check(rejected, "invalid setting rejected");
            Check(((string)settings.GetProperty("PathForUser", All).GetValue(null, null)).Contains(asm.GetName().Name), "host separated user setting");

            Type docs = asm.GetType(prefix + "AgentDocumentation", true), formType = asm.GetType(prefix + "AgentDocumentForm", true);
            foreach (bool guide in new bool[] { false, true })
            {
                string full = (string)docs.GetMethod("Read", All).Invoke(null, new object[] { guide });
                string project = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetFullPath(args[0])));
                string expected = File.ReadAllText(Path.Combine(project, "build", guide ? "ComposedUserGuide.md" : "ComposedReadme.md"), Encoding.UTF8);
                
                // Match the compiler's universal-newline UTF-8 input normalization.
                expected = expected.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", Environment.NewLine);
                Check(full == expected, "normalized embedded document matches this project excluding delivery evidence");
                string overview = (string)docs.GetMethod("Overview", All).Invoke(null, new object[] { full });
                // Only the developer log promises the Agent extension wording; the user guide overview is its own install section.
                Check((guide || overview.Contains("Agent扩展")) && overview.Length < full.Length, "overview extracted from the first chapter");
                using (Form form = (Form)Activator.CreateInstance(formType, All, null, new object[] { guide }, null))
                {
                    Check(!form.Modal && form.ClientSize == new Size(920, 680), "nonmodal initial reader size");
                    Check((string)formType.GetProperty("CurrentSource", All).GetValue(form, null) == overview, "initial overview");
                    Check(form.Text.StartsWith(guide ? "用户预览" : "开发预览", StringComparison.Ordinal), "renamed reader title");
                    int authorizations = 0;
                    formType.GetField("AuthorizeFullText", All).SetValue(form, new Func<bool>(delegate { authorizations++; return true; }));
                    formType.GetField("ReportError", All).SetValue(form, new Action<string>(delegate(string text) { throw new Exception("Unexpected reader error: " + text); }));
                    var selectedFiles = new List<string>();
                    files.GetField("OpenWithDialog", All).SetValue(null, new Func<IntPtr, string, int>(delegate(IntPtr owner, string file) { selectedFiles.Add(file); return selectedFiles.Count == 1 ? 0 : unchecked((int)0x800704C7); }));
                    foreach (Control c in Controls(form)) if (c is Button && c.Text == "全文")
                        c.GetType().GetMethod("OnClick", All).Invoke(c, new object[] { EventArgs.Empty });
                    Call(form, "SetView", true);
                    Check(selectedFiles.Count == 2, "each full click invokes chooser even on cancellation");
                    Check(authorizations == (guide ? 0 : 2), "developer authorizes every full click; guide never authorizes");
                    if (!guide)
                    {
                        formType.GetField("AuthorizeFullText", All).SetValue(form, new Func<bool>(delegate { authorizations++; return false; }));
                        Call(form, "SetView", true);
                        Check(selectedFiles.Count == 2 && !form.IsDisposed, "authorization cancellation does not launch chooser or close reader");
                        formType.GetField("AuthorizeFullText", All).SetValue(form, new Func<bool>(delegate { authorizations++; return true; }));
                    }
                    Check(selectedFiles[0] == selectedFiles[1] && File.Exists(selectedFiles[0]), "stable export persists after application choice");
                    Check(File.ReadAllText(selectedFiles[0], Encoding.UTF8) == full && Path.GetExtension(selectedFiles[0]) == ".md", "UTF8 Markdown exported verbatim");
                    Check(Path.GetFileName(selectedFiles[0]) == (guide ? "A-UserGuide.md" : "B-Devlog.md") && selectedFiles[0].Contains(asm.GetName().Name), "document and project separated export");
                    Check((string)formType.GetProperty("CurrentSource", All).GetValue(form, null) == overview, "full click retains overview");
                    string failure = null;
                    formType.GetField("ReportError", All).SetValue(form, new Action<string>(delegate(string text) { failure = text; }));
                    files.GetField("OpenWithDialog", All).SetValue(null, new Func<IntPtr, string, int>(delegate(IntPtr owner, string file) { return unchecked((int)0x80004005); }));
                    Call(form, "SetView", true);
                    Check(failure != null && failure.Contains(selectedFiles[0]), "chooser failure includes exported path");
                    failure = null;
                    string blocked = Path.Combine(temporary, "blocked"); File.WriteAllText(blocked, "not a directory");
                    files.GetField("CacheRoot", All).SetValue(null, blocked);
                    Call(form, "SetView", true);
                    Check(failure != null && !form.IsDisposed, "export failure keeps reader available");
                    files.GetField("CacheRoot", All).SetValue(null, exportRoot);
                    RichTextBox box = null; foreach (Control c in Controls(form)) if (c is RichTextBox) box = (RichTextBox)c;
                    Check(box != null && box.ReadOnly && box.Text.Contains("Agent-Evo") && box.ScrollBars == RichTextBoxScrollBars.Both, "Chinese readable scroll and selectable copy");
                    box.Select(0, Math.Min(30, box.TextLength)); Check(box.SelectedText.Length > 0, "text selection");
                    Call(form, "SetView", false); Check(box.Text.Length > 0, "return to overview");
                    Check(authorizations == (guide ? 0 : 5), "overview never authorizes; developer has no cached authorization");
                    CheckOverviewFormatting(form, box, overview);
                    string copied = null;
                    formType.GetField("CopyText", All).SetValue(form, new Action<string>(delegate(string text) { copied = text; }));
                    foreach (Control c in Controls(form)) if (c is Button && c.Text == "复制内容")
                        c.GetType().GetMethod("OnClick", All).Invoke(c, new object[] { EventArgs.Empty });
                    Check(copied == box.Text, "copy button passes displayed contents");
                    Check(form.CancelButton != null && ((Control)form.CancelButton).Text == "关闭", "Esc closes reader");
                }
            }
            foreach (string bad in new string[] { "", "**标题**\n", "# 一. 甲\n", "# 一. 甲\n\n## 1. 乙\n" })
            { bool rejectedDoc = false; try { docs.GetMethod("Overview", All).Invoke(null, new object[] { bad }); } catch (TargetInvocationException) { rejectedDoc = true; } Check(rejectedDoc, "empty overview rejected"); }
            CheckPasswordForm(asm);
            Type windows = asm.GetType(prefix + "AgentDocumentWindows", true);
            Form opened = (Form)windows.GetMethod("Open", All).Invoke(null, new object[] { false });
            Check(ReferenceEquals(opened, windows.GetMethod("Open", All).Invoke(null, new object[] { false })), "repeat opens activate existing reader");
            object[] keys = { new Message(), Keys.Escape };
            Check((bool)opened.GetType().GetMethod("ProcessCmdKey", All).Invoke(opened, keys) && opened.IsDisposed, "Esc actually closes reader");
            Form other = (Form)windows.GetMethod("Open", All).Invoke(null, new object[] { true });
            windows.GetMethod("CloseAll", All).Invoke(null, null); Check(other.IsDisposed, "shutdown closes readers");

            Type connectType = asm.GetType(prefix + "Connect", true), cfg = asm.GetType(prefix + "RibbonConfig", true);
            var urls = new List<string>(); FieldInfo launcher = connectType.GetField("AgentWebsiteLauncher", All);
            launcher.SetValue(null, new Action<string>(delegate(string url) { urls.Add(url); }));
            object connect = Activator.CreateInstance(connectType);
            var callbackState = Activator.CreateInstance(state, All, null, new object[] { new Func<bool>(delegate { return false; }), new Action<bool>(delegate(bool v) { }), log }, null);
            connectType.GetField("agentController", All).SetValue(null, callbackState);
            connectType.GetMethod("OnAgentRibbonLoad").Invoke(connect, new object[] { new FakeAgentRibbon() });
            connectType.GetMethod("OnToggleAgentExtension").Invoke(connect, new object[] { null, true });
            Check((bool)connectType.GetMethod("GetAgentTabVisible").Invoke(connect, new object[] { null }) && (bool)connectType.GetMethod("GetAgentExtensionPressed").Invoke(connect, new object[] { null }), "public visibility and pressed callbacks agree");
            connectType.GetMethod("OnToggleAgentExtension").Invoke(connect, new object[] { null, false });
            Check(!(bool)connectType.GetMethod("GetAgentTabVisible").Invoke(connect, new object[] { null }), "public hide callback");
            var xml = new XmlDocument(); xml.LoadXml((string)cfg.GetProperty("RibbonXml", All).GetValue(null, null));
            var ns = new XmlNamespaceManager(xml.NameTable); ns.AddNamespace("r", xml.DocumentElement.NamespaceURI);
            Check(((XmlElement)xml.SelectSingleNode("//r:button[@id='btnAgentReadme']", ns)).GetAttribute("label") == "开发预览", "development menu label");
            Check(((XmlElement)xml.SelectSingleNode("//r:button[@id='btnAgentUserGuide']", ns)).GetAttribute("label") == "用户预览", "user menu label");
            Check(xml.SelectNodes("//r:tab", ns).Count == 2, "two tabs in one add-in");
            Check(((XmlElement)xml.SelectSingleNode("//r:group[@id='grpAgentResources']", ns)).GetAttribute("label")=="详情", "Agent details group");
            Check(((XmlElement)xml.SelectSingleNode("//r:group[@id='grpAgentOperations']", ns)).GetAttribute("label")=="操作", "Agent operations group");
            Check(((XmlElement)xml.SelectSingleNode("//r:button[@id='btnAgentDialogSizes']", ns)).GetAttribute("onAction")=="OnOpenAgentDialogSizes", "public customization callback");
            Check(xml.DocumentElement.GetAttribute("onLoad") == "OnAgentRibbonLoad", "load callback");
            foreach (XmlElement menu in xml.SelectNodes("//r:tab[@id='tabAgentEvo']//r:menu", ns)) Check(!menu.HasAttribute("onAction"), "expansion and Esc never launch website");
            Check(urls.Count == 0, "no website opened while reading menus");
            Check(xml.SelectNodes("//r:menu[@id='menuAgentApi']/r:menu", ns).Count == 2, "two submenu branches");
            Check(xml.SelectNodes("//r:menu[@id='menuAgentCloud']/r:button", ns).Count == 2, "two cloud api leaves");
            Check(xml.SelectNodes("//r:menu[@id='menuAgentSites']/r:menu", ns).Count == 2, "china and overseas branches");
            Check(xml.SelectNodes("//r:menu[@id='menuAgentChina']/r:button", ns).Count == 6, "six domestic agents");
            Check(xml.SelectNodes("//r:menu[@id='menuAgentOverseas']/r:button", ns).Count == 3, "three overseas agents");
            foreach (XmlElement leaf in xml.SelectNodes("//r:menu[@id='menuAgentApi']//r:button", ns))
            {
                string action = leaf.GetAttribute("onAction"); connectType.GetMethod(action).Invoke(connect, new object[] { null });
                Check(urls[urls.Count - 1] == (string)cfg.GetMethod("GetUrl", All).Invoke(null, new object[] { leaf.GetAttribute("id") }), "dedicated URL dispatch " + leaf.GetAttribute("id"));
            }
            Check(urls.Count == 11, "eleven configured urls");
            string home = null;
            files.GetField("BrowserLauncher", All).SetValue(null, new Action<string>(delegate(string pathToHome) { home = pathToHome; }));
            connectType.GetMethod("OnAgentHome").Invoke(connect, new object[] { null });
            string projectRoot = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetFullPath(args[0])));
            string websitePath = args.Length > 1 ? Path.GetFullPath(args[1]) : Path.Combine(Directory.GetParent(projectRoot).FullName, "C-03-Website·Xycabin", "AA-Program", "XyCabin.html");
            Check(home != null && File.ReadAllText(home, Encoding.UTF8) == CorePage(File.ReadAllText(websitePath, Encoding.UTF8)), "home callback opens complete embedded site");
            Check(home.Contains(Path.Combine("Xy", "AgentEvo", "Home")) && Path.GetFileName(home) == "index.html", "home independent of install and shadow-cache paths");
            connectType.GetMethod("OnAgentHome").Invoke(connect, new object[] { null });
            Check(home == (string)files.GetMethod("ExportHome", All).Invoke(null, null), "same content uses same cached website");
            File.WriteAllText(home, "damaged");
            files.GetMethod("ExportHome", All).Invoke(null, null);
            Check(File.ReadAllText(home, Encoding.UTF8) == CorePage(File.ReadAllText(websitePath, Encoding.UTF8)), "damaged exported page repaired with configured content");
            Check(File.Exists((string)files.GetMethod("DefaultBrowserExecutable", All).Invoke(null, null)), "system URL association resolves installed browser");
            Console.WriteLine("AgentExtensionProbe: " + passed + " PASS, 0 FAIL"); return 0;
        }
        catch (Exception ex) { Console.WriteLine("FAIL: " + ex); return 1; }
        finally
        {
            // Only the unique test directory created above, never an arbitrary temporary root.
            if (temporary != null && Directory.Exists(temporary))
            {
                string testRoot = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                if (!Path.GetFullPath(temporary).StartsWith(testRoot, StringComparison.OrdinalIgnoreCase) || !Path.GetFileName(temporary).StartsWith("agent-probe-", StringComparison.Ordinal))
                    throw new InvalidOperationException("Probe cleanup path is outside the test temporary directory");
                Directory.Delete(temporary, true);
            }
        }
    }
}
