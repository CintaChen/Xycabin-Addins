using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // Sequential STA controller. Only private disk copies are ever opened or saved.
    internal sealed class SummaryFileBatch : IDisposable
    {
        private readonly object app;
        private readonly Form dialog = new Form();
        private readonly TableLayoutPanel settings = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3 };
        private readonly TextBox input = new TextBox { Dock = DockStyle.Fill, ReadOnly = true }, output = new TextBox { Dock = DockStyle.Fill }, presetPath = new TextBox { Dock = DockStyle.Fill, ReadOnly = true };
        private readonly TextBox pathSheet = new TextBox { Text = "单体数据汇总", Dock = DockStyle.Fill }, pathCell = new TextBox { Text = "A1", Dock = DockStyle.Fill };
        private readonly DataGridView files = new ResultGrid { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false };
        private readonly RichTextBox log = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true };
        private readonly ProgressBar progress = new ProgressBar { Dock = DockStyle.Top };
        private readonly RoundedButton preview = Ui.Button("批量预检", 120), execute = Ui.Button("执行并另存", 135), stop = Ui.Button("停止", 90);
        private readonly Timer timer = new Timer { Interval = 40 };
        private List<DataGridViewRow> queue;
        private BatchPreset preset;
        private string destination, sourceSheet, sourceCell, scratch, workingFile;
        private int position;
        private bool writing, stopping, inTick;
        private ComBag refs;
        private object workBook;
        private SingleSummaryForm workerForm;
        private SummaryAreaHost worker;
        internal static void Show(object application, Form owner)
        {
            using (var batch = new SummaryFileBatch(application)) batch.dialog.ShowDialog(owner);
        }
        private SummaryFileBatch(object application)
        {
            app = application; dialog.Text = "合并批量链接 · 多文件";
            DialogSizing.Prepare(dialog, "summary.single", "table", Ui.Blue);
            settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var chooseInput = Ui.Button("选择文件夹…", 140);
            chooseInput.Click += delegate { Attempt(delegate {
                string path = HostDialogService.Folder(dialog, "选择待处理工作簿文件夹", false); if (string.IsNullOrEmpty(path)) return;
                var found = Directory.GetFiles(path).Where(IsWorkbook).Where(p => !Path.GetFileName(p).StartsWith("~$", StringComparison.Ordinal)).OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray();
                files.Rows.Clear(); foreach (string file in found) { int row = files.Rows.Add(DefaultSelected(file), Path.GetFileName(file), "待预检"); files.Rows[row].Tag = file; } input.Text = path;
                Message("已列出 " + found.Length + " 个工作簿，默认勾选文件名含（合并）或(合并)的合并主体；可手动调整。");
            }); };
            var choosePreset = Ui.Button("选择导出预设…", 140);
            choosePreset.Click += delegate { Attempt(delegate { string[] paths = HostDialogService.OpenFiles(dialog, "合并链接预设 (*.xybatch.json)|*.xybatch.json"); if (paths.Length != 1) return; BatchPresetStore.Decode(File.ReadAllText(paths[0]), "summary"); presetPath.Text = paths[0]; }); };
            var chooseOutput = Ui.Button("选择输出文件夹…", 155);
            chooseOutput.Click += delegate { string path = HostDialogService.Folder(dialog, "选择结果输出文件夹", true); if (!string.IsNullOrEmpty(path)) output.Text = path; };
            AddRow("输入文件夹", input, chooseInput); AddRow("使用预设", presetPath, choosePreset); AddRow("输出文件夹", output, chooseOutput);
            AddRow("各文件路径来源表", pathSheet, new Label { AutoSize = true, Text = "使用各文件自己的 Lead、变量和期间" }); AddRow("路径来源单元格", pathCell, new Label { AutoSize = true, Text = "默认 A1，读取路径文字或公式结果" });
            files.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "处理", Width = 60 });
            files.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "文件", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            files.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "状态", ReadOnly = true, Width = 200 }); Ui.StyleGrid(files);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(16) };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 50)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 50)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var actions = Ui.ActionBar(); actions.Controls.AddRange(new Control[] { preview, execute, stop });
            root.Controls.Add(settings, 0, 0); root.Controls.Add(files, 0, 1); root.Controls.Add(progress, 0, 2); root.Controls.Add(log, 0, 3); root.Controls.Add(actions, 0, 4); dialog.Controls.Add(root);
            preview.Click += delegate { Attempt(delegate { Begin(false); }); }; execute.Click += delegate { Attempt(delegate { Begin(true); }); };
            stop.Enabled = false; stop.Click += delegate { stopping = true; Message("正在停止；当前临时副本将丢弃，已经输出的成功文件保留。"); };
            timer.Tick += delegate { Advance(); };
            dialog.FormClosing += delegate(object sender, FormClosingEventArgs e) { if (queue != null) { e.Cancel = true; stopping = true; Message("正在停止，请等待当前操作结束后关闭。"); } };
        }
        private void AddRow(string label, Control value, Control action)
        { int row = settings.RowCount++; settings.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left }, 0, row); settings.Controls.Add(value, 1, row); settings.Controls.Add(action, 2, row); }
        internal static bool IsWorkbook(string file)
        { string extension = Path.GetExtension(file); return string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) || string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase); }
        internal static bool DefaultSelected(string file)
        {
            string name = Path.GetFileNameWithoutExtension(file);
            return name.Contains("（合并）") || name.Contains("(合并)");
        }
        internal static string CellPath(object value)
        {
            string text = value as string;
            if (string.IsNullOrWhiteSpace(text)) throw new InvalidOperationException("路径来源单元格为空、错误或不是文字。");
            text = text.Trim(); if (text.Length > 1 && text[0] == '"' && text[text.Length - 1] == '"') text = text.Substring(1, text.Length - 2);
            if (!Path.IsPathRooted(text) || !Directory.Exists(text)) throw new InvalidOperationException("路径来源单元格中的文件夹不存在或不是完整路径：" + text);
            return text;
        }
        internal static BatchPreset Rebind(BatchPreset original, string file, string folder)
        { var copy = BatchPresetStore.Decode(BatchPresetStore.Encode(original), "summary"); copy.Context = file; copy.Values["projectFolder"] = folder; return copy; }
        private void Message(string text) { log.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + text + Environment.NewLine); log.SelectionStart = log.TextLength; log.ScrollToCaret(); }
        private void Attempt(Action action) { try { action(); } catch (Exception error) { Message("失败：" + error.Message); } }
        private void Begin(bool save)
        {
            if (queue != null) return; files.EndEdit();
            var selected = files.Rows.Cast<DataGridViewRow>().Where(r => object.Equals(r.Cells[0].Value, true)).ToList();
            if (selected.Count == 0) throw new InvalidOperationException("请勾选需要处理的文件。");
            var selectedPreset = BatchPresetStore.Decode(File.ReadAllText(presetPath.Text), "summary"); SummaryAreaSettings.Read(selectedPreset);
            string directory = Path.GetFullPath(output.Text.Trim()); if (!Directory.Exists(directory)) throw new InvalidOperationException("请选择已存在的输出文件夹。");
            if (!System.Text.RegularExpressions.Regex.IsMatch(pathCell.Text.Trim(), @"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}$") || string.IsNullOrWhiteSpace(pathSheet.Text)) throw new InvalidOperationException("请填写路径来源表和一个单元格地址，例如 A1。");
            if (save && UiDialog.Show(dialog, "将逐文件重新预检，并覆盖各临时副本中预设指定的目标区域及表头；成功后自动保存到输出文件夹，原文件保持不变，同名输出不覆盖。失败文件跳过，其他文件继续。确认执行？", "批量输出确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.OK) return;
            preset = selectedPreset; destination = directory; sourceSheet = pathSheet.Text.Trim(); sourceCell = pathCell.Text.Trim(); writing = save; stopping = false; position = 0; queue = selected;
            foreach (var row in queue) row.Cells[2].Value = "等待";
            settings.Enabled = files.Enabled = preview.Enabled = execute.Enabled = false; stop.Enabled = true; progress.Maximum = queue.Count; progress.Value = 0;
            Message((save ? "开始执行" : "开始只读预检") + "，共 " + queue.Count + " 个文件；每个文件重新读取自己的数据，预检不保存任何工作簿。"); timer.Start();
        }
        private void OpenCurrent()
        {
            string inputFile = Convert.ToString(queue[position].Tag);
            string outputFile = Path.Combine(destination, Path.GetFileName(inputFile));
            if (string.Equals(Path.GetFullPath(inputFile), Path.GetFullPath(outputFile), StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("输出文件夹不能覆盖输入文件，请选择其他文件夹。");
            if (File.Exists(outputFile)) throw new InvalidOperationException("输出文件已存在，已跳过，不覆盖。");
            using (var bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)app).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++)
                {
                    dynamic existing = bag.Own(books.Item(i));
                    if (string.Equals(Convert.ToString(existing.Name), Path.GetFileName(inputFile), StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("同名工作簿已打开，请先自行保存并关闭后重试。");
                }
            }
            scratch = Path.Combine(destination, ".xybatch-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(scratch);
            workingFile = Path.Combine(scratch, Path.GetFileName(inputFile));
            using (var source = new FileStream(inputFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var target = new FileStream(workingFile, FileMode.CreateNew, FileAccess.Write)) source.CopyTo(target);
            refs = new ComBag();
            using (var quiet = new CompanyWorkbookQuiet(app))
            {
                dynamic books = refs.Own(((dynamic)app).Workbooks); workBook = refs.Own(books.Open(workingFile, 0, false));
                if (workBook == null) throw new InvalidOperationException("无法打开工作副本。");
                MigrationWindowScope.SetOwnedVisible(workBook, false);
            }
            string folder;
            using (var bag = new ComBag())
            {
                dynamic sheets = bag.Own(((dynamic)workBook).Worksheets), sheet = bag.Own(sheets.Item(sourceSheet)), cell = bag.Own(sheet.Range[sourceCell]);
                folder = CellPath((object)cell.Value2);
            }
            workerForm = new SingleSummaryForm(true); var view = (SummaryAreaView)workerForm.Controls[0];
            view.LogSink = delegate(string text) { Message(Path.GetFileName(inputFile) + "｜" + text); };
            worker = new SummaryAreaHost(app, workerForm, view, null, delegate { return true; });
            worker.PrepareBatchFile(Rebind(preset, workingFile, folder));
            if (writing) worker.StartBatchFile();
        }
        private void Advance()
        {
            if (inTick || queue == null) return; inTick = true; timer.Stop();
            try
            {
                if (stopping) { Cleanup(); End(); return; }
                if (position >= queue.Count) { End(); return; }
                try
                {
                    if (worker == null)
                    {
                        queue[position].Cells[2].Value = "预检中"; OpenCurrent();
                        if (writing) { queue[position].Cells[2].Value = "执行中"; return; }
                    }
                    else if (writing && !worker.StepBatchFile()) return;
                    if (writing)
                    {
                        worker.FinishBatchFile();
                        using (var quiet = new CompanyWorkbookQuiet(app)) ((dynamic)workBook).Save();
                        string temporary = workingFile, final = Path.Combine(destination, Path.GetFileName(Convert.ToString(queue[position].Tag)));
                        CloseObjects(); File.Move(temporary, final);
                        Message("已输出：" + final);
                    }
                    queue[position].Cells[2].Value = writing ? "已输出" : "预检通过";
                }
                catch (Exception error) { queue[position].Cells[2].Value = "失败：" + error.Message; Message(Convert.ToString(queue[position].Tag) + "｜失败：" + error.Message); }
                Cleanup(); position++; progress.Value = position;
            }
            catch (Exception error) { Message("批处理停止：" + error.Message); End(); }
            finally { inTick = false; if (queue != null) timer.Start(); }
        }
        private void End()
        { timer.Stop(); queue = null; settings.Enabled = files.Enabled = preview.Enabled = execute.Enabled = true; stop.Enabled = false; Message("本轮结束；成功输出保留，失败原因见文件列表及日志。原文件未修改。"); }
        private void CloseObjects()
        {
            if (worker != null) { worker.Dispose(); worker = null; }
            if (workerForm != null) { workerForm.Dispose(); workerForm = null; }
            if (workBook != null) { using (var quiet = new CompanyWorkbookQuiet(app)) { quiet.ExcludeClosingBook(workBook); ((dynamic)workBook).Close(false); } workBook = null; }
            if (refs != null) { refs.Dispose(); refs = null; }
        }
        private void Cleanup()
        {
            CloseObjects();
            if (workingFile != null && File.Exists(workingFile)) File.Delete(workingFile);
            if (scratch != null && Directory.Exists(scratch)) Directory.Delete(scratch, false);
            workingFile = scratch = null;
        }
        public void Dispose() { timer.Stop(); timer.Dispose(); try { Cleanup(); } finally { dialog.Dispose(); } }
    }
}
