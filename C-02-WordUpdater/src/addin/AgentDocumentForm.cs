using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    internal static class AgentDocumentation
    {
        internal static string Read(bool userGuide)
        {
            string name = userGuide ? "WpsShared.Agent.UserGuide.md" : "WpsShared.Agent.Readme.md";
            using (Stream stream = typeof(Connect).Assembly.GetManifestResourceStream(name))
            {
                if (stream == null) throw new FileNotFoundException("缺少内嵌文档 " + name);
                using (var reader = new StreamReader(stream, Encoding.UTF8, true)) return reader.ReadToEnd();
            }
        }
        private static void SkipBlank(string[] lines, ref int index)
        {
            while (index < lines.Length && lines[index].Trim().Length == 0) index++;
        }
        private static string Paragraph(string[] lines, ref int index)
        {
            var text = new StringBuilder();
            while (index < lines.Length)
            {
                string line = lines[index].Trim();
                if (line.Length == 0 || line.StartsWith("#", StringComparison.Ordinal)) break;
                if (text.Length > 0) text.Append('\n');
                text.Append(line);
                index++;
            }
            return text.ToString().Trim();
        }
        internal static string Overview(string source)
        {
            // The overview is the first paragraph after the document title/date, or
            // after the first level-one chapter. No invisible markers are required.
            string[] lines = (source ?? "").Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            int index = 0;
            SkipBlank(lines, ref index);
            if (index < lines.Length)
            {
                string first = lines[index].Trim();
                if (first.StartsWith("**", StringComparison.Ordinal) && first.EndsWith("**", StringComparison.Ordinal))
                {
                    index++;
                    SkipBlank(lines, ref index);
                }
            }
            if (index < lines.Length && lines[index].Trim().StartsWith("最后更新：", StringComparison.Ordinal))
            {
                index++;
                SkipBlank(lines, ref index);
            }
            if (index < lines.Length && !lines[index].TrimStart().StartsWith("#", StringComparison.Ordinal))
            {
                string text = Paragraph(lines, ref index);
                if (text.Length > 0) return text;
            }
            while (index < lines.Length && (lines[index].Trim().Length == 0 || lines[index].TrimStart().StartsWith("#", StringComparison.Ordinal))) index++;
            string result = Paragraph(lines, ref index);
            if (result.Length == 0) throw new InvalidDataException("文档概览为空");
            return result;
        }
    }

    internal static class AgentDocumentAccess
    {
        private const string PasswordHash = "e4f5461fbec74bd42261d3da6e9df7c057d52abe141761918c959eacf1ed828f";
        internal static bool VerifyPassword(string password)
        {
            byte[] input = Encoding.UTF8.GetBytes((password ?? "").Trim());
            try
            {
                using (var hash = SHA256.Create())
                    return string.Equals(BitConverter.ToString(hash.ComputeHash(input)).Replace("-", "").ToLowerInvariant(), PasswordHash, StringComparison.Ordinal);
            }
            finally { Array.Clear(input, 0, input.Length); }
        }
        // Every developer request asks again. Exceptions/cancellation stop the
        // sequence; no authorization state or entered text is retained here.
        internal static bool OpenFullText(bool userGuide, Func<bool> authorize, Func<string> export, Action<string> launch)
        {
            try
            {
                if (!userGuide && (authorize == null || !authorize())) return false;
                if (export == null || launch == null) return false;
                string path = export();
                if (string.IsNullOrWhiteSpace(path)) return false;
                launch(path); return true;
            }
            catch { return false; }
        }
    }

    internal sealed class DeveloperPasswordForm : Form
    {
        internal readonly TextBox Password = new TextBox { Name = "developerPassword", UseSystemPasswordChar = true, Dock = DockStyle.Fill, MaxLength = 256 };
        internal readonly Label Hint = new Label { Name = "passwordHint", Text = "请输入开发者的微信名称", AutoSize = true };
        internal readonly Label Error = new Label { Name = "passwordError", Text = "", AutoSize = true, ForeColor = Ui.Danger };
        internal Func<string, bool> Verify = AgentDocumentAccess.VerifyPassword;
        internal DeveloperPasswordForm()
        {
            Text = "开发全文验证";
            DialogSizing.Prepare(this, "agent.developerpassword", "settings", Ui.Blue);
            ShowInTaskbar = false; MinimizeBox = false; MaximizeBox = false; KeyPreview = true;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 4 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Hint.Margin = new Padding(0, 0, 0, 12); Password.Margin = new Padding(0, 0, 0, 12); Error.Margin = Padding.Empty;
            Hint.TabIndex = 0; Password.TabIndex = 1; Error.TabIndex = 2;
            var actions = Ui.ActionBar(); actions.FlowDirection = FlowDirection.RightToLeft; actions.Margin = Padding.Empty; actions.TabIndex = 3;
            var cancel = Ui.Button("取消", 100); cancel.Name = "passwordCancel"; cancel.DialogResult = DialogResult.Cancel;
            var confirm = Ui.Button("确认", 100); confirm.Name = "passwordConfirm"; Ui.StylePrimary(confirm);
            confirm.Click += delegate { Submit(); }; cancel.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
            AcceptButton = confirm; CancelButton = cancel; actions.Controls.Add(cancel); actions.Controls.Add(confirm);
            layout.Controls.Add(Hint, 0, 0); layout.Controls.Add(Password, 0, 1); layout.Controls.Add(Error, 0, 2); layout.Controls.Add(actions, 0, 3);
            Controls.Add(layout);
        }
        internal void Submit()
        {
            if (DialogPreviewFactory.IsConstructing || DialogPreviewFactory.IsPreview(this)) return;
            bool accepted = false;
            try { accepted = Verify != null && Verify(Password.Text); }
            catch { Error.Text = "验证未完成，请重试。"; Password.Clear(); if (Visible) Password.Focus(); return; }
            Password.Clear();
            if (!accepted) { Error.Text = "名称不正确，请重试。"; if (Visible) Password.Focus(); return; }
            DialogResult = DialogResult.OK; Close();
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); return true; }
            if (keyData == Keys.Enter) { Submit(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            Password.Clear();
            if (DialogResult != DialogResult.OK) DialogResult = DialogResult.Cancel;
            base.OnFormClosing(e);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { Password.Clear(); Verify = null; }
            base.Dispose(disposing);
        }
    }

    internal static class AgentDocumentWindows
    {
        private static readonly Dictionary<bool, AgentDocumentForm> windows = new Dictionary<bool, AgentDocumentForm>();
        internal static AgentDocumentForm Open(bool userGuide)
        {
            AgentDocumentForm form;
            if (windows.TryGetValue(userGuide, out form) && !form.IsDisposed)
            {
                if (form.WindowState == FormWindowState.Minimized) form.WindowState = FormWindowState.Normal;
                form.Activate(); return form;
            }
            form = new AgentDocumentForm(userGuide); windows[userGuide] = form;
            form.FormClosed += delegate { windows.Remove(userGuide); };
            form.Show(); return form;
        }
        internal static void CloseAll()
        {
            foreach (AgentDocumentForm form in new List<AgentDocumentForm>(windows.Values)) form.Close();
            windows.Clear();
        }
    }

    internal sealed class AgentDocumentForm : Form
    {
        private readonly string full, overview;
        private readonly bool userGuide;
        private readonly RichTextBox document;
        private readonly RoundedButton overviewButton, fullButton;
        internal Action<string> CopyText = delegate(string text) { Clipboard.SetText(text); };
        internal Action<string> OpenFullText;
        internal Func<bool> AuthorizeFullText;
        internal Action<string> ReportError = delegate(string message) { UiDialog.Show(message, "开发流程", MessageBoxButtons.OK, MessageBoxIcon.Warning); };
        internal string CurrentSource { get { return overview; } }

        internal AgentDocumentForm(bool userGuide)
        {
            this.userGuide = userGuide;
            OpenFullText = delegate(string path) { AgentLocalFiles.OpenDocument(Handle, path); };
            AuthorizeFullText = delegate { return HostDialogService.ShowIsolatedResult(this, delegate { return new DeveloperPasswordForm(); }) == DialogResult.OK; };
            full = AgentDocumentation.Read(userGuide); overview = AgentDocumentation.Overview(full);
            Text = (userGuide ? "用户预览" : "开发预览") + " · " + typeof(Connect).Assembly.GetName().Name;
            DialogSizing.Prepare(this, userGuide ? "agent.user" : "agent.developer", "document", Ui.Blue);
            KeyPreview = true;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 3 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var toolbar = Ui.ActionBar(); toolbar.Margin = new Padding(0, 0, 0, 12);
            overviewButton = Ui.Button("概览", 100); fullButton = Ui.Button("全文", 100);
            toolbar.TabIndex = 0; overviewButton.TabIndex = 0; fullButton.TabIndex = 1;
            overviewButton.Click += delegate { SetView(false); }; fullButton.Click += delegate { SetView(true); };
            toolbar.Controls.Add(overviewButton); toolbar.Controls.Add(fullButton);
            fullButton.AccessibleDescription = "选择电脑中的应用打开 Markdown 全文";
            var help = new Label { Text = "全文：选择应用打开 Markdown 文件", AutoSize = true,
                ForeColor = Ui.InkDim, Font = Ui.MakeFont(9, FontStyle.Regular, null), Margin = new Padding(12, 11, 0, 0) };
            toolbar.Controls.Add(help);
            document = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BorderStyle = BorderStyle.None,
                BackColor = Color.White, ForeColor = Ui.Ink, Font = Ui.MakeFont(10, FontStyle.Regular, null),
                ScrollBars = RichTextBoxScrollBars.Both, DetectUrls = false, WordWrap = true, HideSelection = false,
                AccessibleName = "文档内容", TabIndex = 2 };
            var inner = new Panel { Dock = DockStyle.Fill, Padding = new Padding(14), BackColor = Color.White };
            inner.Controls.Add(document);
            var card = Ui.Card(inner, ""); card.Margin = new Padding(0, 0, 0, 12);
            card.TabIndex = 1; inner.TabIndex = 0; document.TabIndex = 0;
            var actions = Ui.ActionBar(); actions.FlowDirection = FlowDirection.RightToLeft; actions.Margin = Padding.Empty;
            actions.TabIndex = 2;
            var close = Ui.Button("关闭", 100); close.Click += delegate { Close(); }; CancelButton = close;
            close.TabIndex = 1;
            var copy = Ui.Button("复制内容", 120);
            copy.TabIndex = 0;
            copy.Click += delegate
            {
                try { CopyText(document.Text); }
                catch (Exception ex) { UiDialog.Show("无法复制内容：" + ex.Message, "开发流程", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
            actions.Controls.Add(close); actions.Controls.Add(copy);
            layout.Controls.Add(toolbar, 0, 0); layout.Controls.Add(card, 0, 1); layout.Controls.Add(actions, 0, 2);
            Controls.Add(layout); SetView(false);
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape) { Close(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        internal void SetView(bool fullView)
        {
            if (fullView)
            {
                if (DialogPreviewFactory.IsConstructing || DialogPreviewFactory.IsPreview(this)) return;
                string path = "";
                Exception failure = null; string stage = "验证";
                AgentDocumentAccess.OpenFullText(userGuide,
                    delegate { try { return AuthorizeFullText != null && AuthorizeFullText(); } catch (Exception ex) { failure = ex; return false; } },
                    delegate { stage = "导出"; try { path = AgentLocalFiles.ExportDocument(userGuide, full); return path; } catch (Exception ex) { failure = ex; throw; } },
                    delegate(string exported) { stage = "打开"; try { OpenFullText(exported); } catch (Exception ex) { failure = ex; throw; } });
                if (failure != null)
                {
                    // Never include authorization input or exception messages.
                    Connect.Log("Agent 全文" + stage + "失败：" + failure.GetType().Name);
                    ReportError("无法完成全文" + stage + "，请重试。" + (path.Length == 0 ? "" : "\r\n文件路径：" + path));
                }
                return;
            }
            Ui.StylePrimary(overviewButton); Ui.StyleSecondary(fullButton);
            document.Clear(); bool code = false;
            // Render a safe, readable Markdown subset without a browser or scripts.
            using (Font body = Ui.MakeFont(10, FontStyle.Regular, null), heading = Ui.MakeFont(13, FontStyle.Bold, null),
                mono = new Font(FontFamily.GenericMonospace, 10))
            {
                foreach (string original in CurrentSource.Replace("\r\n", "\n").Split('\n'))
                {
                    string line = original;
                    if (line.TrimStart().StartsWith("```")) { code = !code; continue; }
                    if (!code && line.TrimStart().StartsWith("<!--")) continue;
                    bool title = !code && Regex.IsMatch(line, "^#{1,6}\\s");
                    if (!code)
                    {
                        line = Regex.Replace(line, "^#{1,6}\\s+", "");
                        line = Regex.Replace(line, "\\[([^\\]]+)\\]\\(([^)]+)\\)", "$1（$2）");
                        line = line.Replace("**", "").Replace("`", "");
                    }
                    int start = document.TextLength;
                    document.AppendText(line + "\n");
                    document.Select(start, document.TextLength - start);
                    document.SelectionFont = code ? mono : title ? heading : body;
                    document.SelectionColor = title ? Ui.Blue : Ui.Ink;
                }
            }
            document.Select(0, 0); document.ScrollToCaret();
        }
    }
}
