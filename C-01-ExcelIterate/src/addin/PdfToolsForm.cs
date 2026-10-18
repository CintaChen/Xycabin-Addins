using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace Excel_Iterate
{
    internal sealed class PdfToolRequest
    {
        internal string Mode, Output, Text, Ranges;
        internal string[] Files;
        internal decimal Opacity, Size, Angle;
        internal XmlDocument ToXml()
        {
            var doc = new XmlDocument(); var root = doc.CreateElement("PdfJob"); doc.AppendChild(root);
            root.SetAttribute("mode", Mode); root.SetAttribute("output", Path.GetFullPath(Output));
            root.SetAttribute("text", Text ?? ""); root.SetAttribute("ranges", Ranges ?? "");
            root.SetAttribute("opacity", Opacity.ToString(CultureInfo.InvariantCulture));
            root.SetAttribute("size", Size.ToString(CultureInfo.InvariantCulture)); root.SetAttribute("angle", Angle.ToString(CultureInfo.InvariantCulture));
            foreach (string file in Files) { var child = doc.CreateElement("File"); child.InnerText = Path.GetFullPath(file); root.AppendChild(child); }
            return doc;
        }
    }
    internal static class PdfWorkerBridge
    {
        private static readonly object gate = new object();
        private static string Digest(byte[] data) { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(data)).Replace("-", ""); }
        private static string Extract()
        {
            using (var stream = typeof(PdfWorkerBridge).Assembly.GetManifestResourceStream("Excel_Iterate.PdfWorker.zip"))
            {
                if (stream == null) throw new InvalidOperationException("PDF 工具未打包，请重新构建加载项。");
                byte[] bytes; using (var copy = new MemoryStream()) { stream.CopyTo(copy); bytes = copy.ToArray(); }
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Excel_Iterate", "PdfWorker", Digest(bytes));
                lock (gate)
                {
                    Directory.CreateDirectory(folder);
                    using (var input = new MemoryStream(bytes)) using (var archive = new ZipArchive(input, ZipArchiveMode.Read))
                    foreach (var entry in archive.Entries)
                    {
                        if (entry.FullName != Path.GetFileName(entry.FullName)) throw new InvalidDataException("PDF 工具包含无效路径。");
                        string target = Path.Combine(folder, entry.FullName); byte[] contents;
                        using (var source = entry.Open()) using (var copy = new MemoryStream()) { source.CopyTo(copy); contents = copy.ToArray(); }
                        if (File.Exists(target) && Digest(File.ReadAllBytes(target)) == Digest(contents)) continue;
                        string temp = target + "." + Guid.NewGuid().ToString("N") + ".tmp";
                        try { File.WriteAllBytes(temp, contents); if (File.Exists(target)) File.Replace(temp, target, null); else File.Move(temp, target); }
                        finally { if (File.Exists(temp)) File.Delete(temp); }
                    }
                }
                return Path.Combine(folder, "PdfWorker.exe");
            }
        }
        internal static string Run(PdfToolRequest job, Func<bool> cancelled)
        {
            string executable = Extract();
            string folder = Path.Combine(Path.GetTempPath(), "ExcelIterate-PDF-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
            string request = Path.Combine(folder, "request.xml"); job.ToXml().Save(request);
            try
            {
                if (cancelled()) return "已取消，未开始处理。";
                var start = new ProcessStartInfo(executable, "\"" + request + "\"") { UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = Path.GetDirectoryName(executable) };
                using (var process = Process.Start(start))
                {
                    while (!process.WaitForExit(200)) if (cancelled() && !File.Exists(request + ".cancel")) File.WriteAllText(request + ".cancel", "cancel");
                    if (!File.Exists(request + ".result")) throw new InvalidOperationException("PDF 工具未返回结果，退出码：" + process.ExitCode);
                    var result = new XmlDocument { XmlResolver = null };
                    using (var reader = XmlReader.Create(request + ".result", new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) result.Load(reader);
                    var text = new StringBuilder(); int success = result.SelectNodes("/Result/Success").Count, failed = result.SelectNodes("/Result/Failure").Count;
                    text.AppendLine("已生成 " + success + " 个 PDF，失败 " + failed + " 项。" + (process.ExitCode == 3 ? "已取消。" : ""));
                    foreach (XmlElement line in result.DocumentElement.ChildNodes) text.AppendLine((line.Name == "Success" ? "已保存：" : "提示：") + line.InnerText);
                    return text.ToString();
                }
            }
            finally { try { Directory.Delete(folder, true); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
        }
    }
    internal sealed class PdfToolsForm : Form, IDialogDefaultLayout
    {
        private readonly string mode;
        private readonly ListBox files = new ListBox { Name = "pdfFiles", Dock = DockStyle.Fill, HorizontalScrollbar = true, SelectionMode = SelectionMode.MultiExtended };
        private readonly TextBox output = new TextBox { Name = "pdfOutput", Dock = DockStyle.Fill };
        private readonly TextBox watermark = new TextBox { Name = "watermarkText", Text = "内部资料", Width = 200, MaxLength = 200 };
        private readonly TextBox ranges = new TextBox { Name = "splitRanges", Width = 280 };
        private readonly NumericUpDown opacity = Number("watermarkOpacity", 1, 100, 20);
        private readonly NumericUpDown size = Number("watermarkSize", 8, 144, 48);
        private readonly NumericUpDown angle = Number("watermarkAngle", -180, 180, -35);
        private readonly TextBox log = new TextBox { Name = "pdfLog", Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false };
        private readonly List<Control> editing = new List<Control>();
        private readonly RoundedButton run = Ui.Button("开始处理", 110), cancel = Ui.Button("取消处理", 110);
        private volatile bool cancellation;
        private bool working;
        internal string Mode { get { return mode; } }
        public Size MeasureDefaultClient() { PerformLayout(); return new Size(DialogSizing.LogicalClient(this).Width, 650); }
        internal void SetPreview()
        {
            files.Items.Add(mode == "convert" ? @"C:\示例资料\年度报告.docx" : @"C:\示例资料\年度报告.pdf");
            if (mode == "merge") files.Items.Add(@"C:\示例资料\报告附件.pdf");
            output.Text = @"C:\示例资料\PDF输出";
            if (mode == "split") ranges.Text = "1-3,4-6,7";
            log.Text = "尺寸预览：仅展示示例，不读取或处理文件。\r\n同名输出将自动编号，原文件保留。";
        }
        internal static string Title(string value) { return value == "convert" ? "文件转 PDF" : value == "watermark" ? "PDF 添加水印" : value == "split" ? "PDF 拆分" : "PDF 合并"; }
        private static NumericUpDown Number(string name, int min, int max, int value)
        { return new NumericUpDown { Name = name, Minimum = min, Maximum = max, Value = value, Width = 66 }; }
        internal PdfToolsForm(string operation)
        {
            mode = operation; Text = Title(mode);
            DialogSizing.Prepare(this, "file.pdf." + mode, "document", Ui.Blue);
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 7 };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 60)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 40)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var hint = Ui.Label(mode == "convert" ? "Word／Excel 依照原排版导出；图片适配 A4（含多页 TIFF）。文件全部在本机处理。" : mode == "merge" ? "按列表从上到下合并；使用上移、下移调整顺序。" : mode == "split" ? "页码从 1 开始；留空逐页拆分，或输入 1-3,4-6,7（逗号分隔输出文件）。" : "给每一页叠加居中文字水印，可设置透明度、字号和角度。", 9, false);
            hint.AutoSize = true; hint.Dock = DockStyle.Fill; layout.Controls.Add(hint, 0, 0);
            var toolbar = Ui.ActionBar();
            AddButton(toolbar, "添加文件…", AddFiles); AddButton(toolbar, "移除", delegate { foreach (int index in files.SelectedIndices.Cast<int>().OrderByDescending(i => i).ToArray()) files.Items.RemoveAt(index); });
            AddButton(toolbar, "上移", delegate { MoveFile(-1); }); AddButton(toolbar, "下移", delegate { MoveFile(1); });
            layout.Controls.Add(toolbar, 0, 1); layout.Controls.Add(Ui.Card(files, "来源文件（按列表顺序）"), 0, 2); editing.Add(files);
            var destination = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 3, Padding = new Padding(0, 8, 0, 8) };
            destination.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); destination.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); destination.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var browse = Ui.Button("选择目录…", 110); browse.Click += delegate { string path = HostDialogService.Folder(this, "选择 PDF 输出文件夹", true); if (path != null) output.Text = path; };
            destination.Controls.Add(Ui.Label("输出目录", 10, false), 0, 0); destination.Controls.Add(output, 1, 0); destination.Controls.Add(browse, 2, 0); editing.Add(output); editing.Add(browse);
            layout.Controls.Add(destination, 0, 3);
            var options = Ui.ActionBar(); options.AutoSize = true;
            if (mode == "watermark")
            {
                options.Controls.Add(Ui.Label("文字", 10, false)); options.Controls.Add(watermark);
                options.Controls.Add(Ui.Label("不透明度 %", 10, false)); options.Controls.Add(opacity);
                options.Controls.Add(Ui.Label("字号", 10, false)); options.Controls.Add(size);
                options.Controls.Add(Ui.Label("角度 °", 10, false)); options.Controls.Add(angle);
            }
            if (mode == "split") { options.Controls.Add(Ui.Label("拆分范围", 10, false)); options.Controls.Add(ranges); }
            editing.Add(options); layout.Controls.Add(options, 0, 4); layout.Controls.Add(Ui.Card(log, "处理结果（同名文件自动编号，不覆盖源文件）"), 0, 5);
            var actions = Ui.ActionBar(); Ui.StylePrimary(run); cancel.Enabled = false;
            run.Click += delegate { Start(); }; cancel.Click += delegate { cancellation = true; cancel.Enabled = false; log.AppendText("正在取消，等待当前导出结束…\r\n"); };
            var open = Ui.Button("打开输出目录", 140); open.Click += delegate { try { if (Directory.Exists(output.Text)) Process.Start(new ProcessStartInfo(output.Text) { UseShellExecute = true }); } catch (Exception e) { log.AppendText(e.Message); } };
            actions.Controls.Add(run); actions.Controls.Add(cancel); actions.Controls.Add(open); layout.Controls.Add(actions, 0, 6); Controls.Add(layout);
            FormClosing += delegate(object sender, FormClosingEventArgs e) { if (working) { cancellation = true; if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; log.AppendText("正在取消，请在当前处理结束后关闭窗口。\r\n"); } } };
            Disposed += delegate { cancellation = true; };
            KeyPreview = true; KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) { if (working) cancellation = true; else Close(); e.Handled = true; } };
        }
        private void AddButton(FlowLayoutPanel bar, string text, Action action)
        { var button = Ui.Button(text, 100); button.Click += delegate { action(); }; bar.Controls.Add(button); editing.Add(button); }
        private void AddFiles()
        {
            foreach (string path in HostDialogService.OpenFiles(this, mode == "convert" ? "文档和图片|*.doc;*.docx;*.rtf;*.odt;*.wps;*.xls;*.xlsx;*.xlsm;*.xlsb;*.ods;*.et;*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.gif" : "PDF 文件 (*.pdf)|*.pdf"))
                if (!files.Items.Cast<string>().Contains(path, StringComparer.OrdinalIgnoreCase)) files.Items.Add(path);
            if (output.Text.Length == 0 && files.Items.Count > 0) output.Text = Path.GetDirectoryName((string)files.Items[0]);
        }
        private void MoveFile(int direction)
        {
            if (files.SelectedIndices.Count != 1) return; int index = files.SelectedIndex, target = index + direction;
            if (target < 0 || target >= files.Items.Count) return;
            object value = files.Items[index]; files.Items.RemoveAt(index); files.Items.Insert(target, value); files.SelectedIndex = target;
        }
        private void Start()
        {
            if (working || DialogPreviewFactory.IsPreview(this)) return;
            if (files.Items.Count == 0 || (mode == "merge" && files.Items.Count < 2)) { log.Text = "请添加" + (mode == "merge" ? "至少两个 PDF。" : "来源文件。"); return; }
            if (!Directory.Exists(output.Text.Trim())) { log.Text = "请选择已存在的输出文件夹。"; return; }
            if (mode == "watermark" && string.IsNullOrWhiteSpace(watermark.Text)) { log.Text = "请输入水印文字。"; return; }
            var request = new PdfToolRequest { Mode = mode, Files = files.Items.Cast<string>().ToArray(), Output = output.Text.Trim(), Text = watermark.Text, Ranges = ranges.Text, Opacity = opacity.Value, Size = size.Value, Angle = angle.Value };
            working = true; cancellation = false; foreach (var control in editing) control.Enabled = false; run.Enabled = false; cancel.Enabled = true;
            log.Text = "正在处理，请稍候…\r\n";
            var worker = new BackgroundWorker();
            worker.DoWork += delegate(object sender, DoWorkEventArgs e) { e.Result = PdfWorkerBridge.Run(request, () => cancellation); };
            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                working = false; worker.Dispose(); if (IsDisposed || Disposing) return;
                log.Text = e.Error == null ? Convert.ToString(e.Result) : "处理失败：" + e.Error.Message;
                foreach (var control in editing) control.Enabled = true; run.Enabled = true; cancel.Enabled = false;
            };
            worker.RunWorkerAsync();
        }
    }
}
