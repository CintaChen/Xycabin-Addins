using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// Diagnostic launcher only. Uses the production PDF windows without installing VSTO.
internal static class PdfTrialLauncher
{
    private const BindingFlags All = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
#if WPS_WORD
    private const string Host = "WordUpdater", Product = "Word PDF 集成验证版";
#else
    private const string Host = "Excel_Iterate", Product = "Excel PDF 集成验证版";
#endif
    private static Assembly assembly;
    private static readonly Dictionary<string, Form> forms = new Dictionary<string, Form>();
    private static object FindApplication()
    {
#if WPS_WORD
        string[] ids = { "KWPS.Application", "WPS.Application", "Word.Application" };
#else
        string[] ids = { "KET.Application", "et.Application", "Excel.Application" };
#endif
        foreach (string id in ids) try { return Marshal.GetActiveObject(id); } catch (COMException) { }
        throw new InvalidOperationException("请先在 WPS／Office 中打开需要导出的文档，再点击当前文档入口。");
    }
    private static Form Create(string mode)
    {
        object app = null, session = null; Form form = null; bool transferred = false;
        try
        {
            if (mode == "current")
            {
                app = FindApplication();
                session = Activator.CreateInstance(assembly.GetType(Host + ".PdfCurrentDocument", true), All, null, new[] { app }, null);
            }
            form = (Form)Activator.CreateInstance(assembly.GetType(Host + ".PdfToolsForm", true), All, null, new object[] { mode }, null);
            if (session != null)
            {
                Type type = form.GetType(); FieldInfo export = type.GetField("ExportCurrent", All);
                export.SetValue(form, Delegate.CreateDelegate(export.FieldType, session, session.GetType().GetMethod("Export", All)));
                object ownedApp = app, ownedSession = session;
                type.GetField("CloseCurrent", All).SetValue(form, new Action(delegate { try { ((IDisposable)ownedSession).Dispose(); } finally { if (Marshal.IsComObject(ownedApp)) Marshal.ReleaseComObject(ownedApp); } }));
                transferred = true;
                type.GetMethod("SetCurrentSource", All).Invoke(form, new[] { session.GetType().GetProperty("DisplayName", All).GetValue(session, null) });
            }
            return form;
        }
        catch { if (form != null) form.Dispose(); if (!transferred) { if (session != null) ((IDisposable)session).Dispose(); if (app != null && Marshal.IsComObject(app)) Marshal.ReleaseComObject(app); } throw; }
    }
    private static void Open(string mode)
    {
        Form current; if (forms.TryGetValue(mode, out current) && !current.IsDisposed) { current.Activate(); return; }
        try { var form = Create(mode); forms[mode] = form; form.FormClosed += delegate { forms.Remove(mode); form.Dispose(); }; form.Show(); }
        catch (Exception error) { MessageBox.Show((error.InnerException ?? error).Message, Product); }
    }
    [STAThread] private static int Main(string[] args)
    {
        try
        {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            assembly = Assembly.LoadFrom(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Host + ".dll"));
            if (args.Length > 0 && args[0] == "--check")
            {
                foreach (string mode in new[] { "convert", "watermark", "split", "merge" }) using (Form form = Create(mode)) if (form.Controls.Count == 0) throw new Exception("Empty production form");
                return 0;
            }
            using (var home = new Form { Text = Product, ClientSize = new Size(620, 400), StartPosition = FormStartPosition.CenterScreen, MinimumSize = new Size(580, 420), Font = new Font("Microsoft YaHei UI", 10), BackColor = Color.FromArgb(245, 247, 251) })
            {
                var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), ColumnCount = 2, RowCount = 5 };
                root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42)); for (int i = 0; i < 3; i++) root.RowStyles.Add(new RowStyle(SizeType.Percent, 33.33f)); root.RowStyles.Add(new RowStyle(SizeType.Absolute, 80));
                var title = new Label { Text = "文件转换与 PDF 工具", Dock = DockStyle.Fill, Font = new Font(home.Font, FontStyle.Bold) }; root.Controls.Add(title, 0, 0); root.SetColumnSpan(title, 2);
                string[] modes = { "current", "convert", "watermark", "split", "merge" }, labels = { "当前文档转 PDF", "批量文件转 PDF", "PDF 添加水印", "PDF 拆分", "PDF 合并" };
                for (int i = 0; i < modes.Length; i++) { string mode = modes[i]; var button = new Button { Text = labels[i], Dock = DockStyle.Fill, Margin = new Padding(5), BackColor = Color.White, FlatStyle = FlatStyle.Flat }; button.Click += delegate { Open(mode); }; root.Controls.Add(button, i % 2, i / 2 + 1); }
                var note = new Label { Text = "独立集成验证程序，不是 WPS 安装包。\r\n当前文档入口连接已打开的对应宿主，并在导出前保存修改。\r\n每次执行选择输出目录；新文件输出，同名自动编号。", Dock = DockStyle.Fill, AutoSize = false, TextAlign = ContentAlignment.MiddleLeft }; root.Controls.Add(note, 0, 4); root.SetColumnSpan(note, 2); home.Controls.Add(root);
                home.FormClosed += delegate { foreach (var form in new List<Form>(forms.Values)) form.Dispose(); forms.Clear(); };
                Application.Run(home);
            }
            return 0;
        }
        catch (Exception error) { if (args.Length == 0) MessageBox.Show((error.InnerException ?? error).Message, Product); return 1; }
    }
}
