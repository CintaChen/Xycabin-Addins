using System;
using System.Collections.Generic;
using System.Drawing;
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
        private readonly TableLayoutPanel settings = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 4, Margin = new Padding(0, 0, 0, 12) };
        private readonly TextBox input = new TextBox { Dock = DockStyle.Fill, ShortcutsEnabled = true, AccessibleDescription = "输入或粘贴文件夹路径，按 Enter 或点击读取文件。" }, output = new TextBox { Dock = DockStyle.Fill }, presetPath = new TextBox { Dock = DockStyle.Fill, ReadOnly = true };
        private readonly TextBox pathSheet = new TextBox { Text = "单体数据汇总", Dock = DockStyle.Fill }, pathCell = new TextBox { Text = "A1", Dock = DockStyle.Fill };
        private readonly DataGridView files = new ResultGrid { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false };
        private readonly RichTextBox log = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true };
        private readonly ProgressBar progress = new ProgressBar { Dock = DockStyle.Top };
        private readonly RoundedButton preview = Ui.Button("批量预检", 120), execute = Ui.Button("执行并另存", 135), stop = Ui.Button("停止", 90);
        private readonly Timer timer = new Timer { Interval = 40 };
        private List<DataGridViewRow> queue;
        private BatchPreset preset;
        private BatchPreset chosenPreset;
        private List<CompanyStructureTask> structureTasks;
        internal bool StructureUsesLevel = true;
        private string destination, sourceSheet, sourceCell, scratch, workingFile, loadedInput;
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
            settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var chooseInput = Ui.Button("选择文件夹", 155);
            chooseInput.Click += delegate { Attempt(delegate {
                string path = HostDialogService.Folder(dialog, "选择待处理工作簿文件夹", false); if (string.IsNullOrEmpty(path)) return;
                input.Text = path; LoadFiles();
            }); };
            var loadInput = Ui.Button("读取文件", 115); loadInput.Click += delegate { Attempt(LoadFiles); };
            input.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Attempt(LoadFiles); } };
            input.TextChanged += delegate { if (queue == null) { loadedInput = null; files.Rows.Clear(); } };
            var choosePreset = Ui.Button("选择已保存预设", 155);
            choosePreset.Click += delegate { Attempt(delegate {
                string json = HostDialogService.ShowIsolatedResult<string>(dialog,
                    delegate { return new BatchPresetManagerForm("summary"); },
                    delegate(Form chooser) { var value = ((BatchPresetManagerForm)chooser).Chosen; return value == null ? null : BatchPresetStore.Encode(value); }, null);
                if (json != null) UsePreset(BatchPresetStore.Decode(json, "summary"), "已保存：");
            }); };
            var importPreset = Ui.Button("导入预设", 115);
            importPreset.Click += delegate { Attempt(delegate { string[] paths = HostDialogService.OpenFiles(dialog, "合并链接预设 (*.xybatch.json)|*.xybatch.json"); if (paths.Length != 1) return; UsePreset(BatchPresetStore.Read(paths[0], "summary"), "已导入："); }); };
            var chooseOutput = Ui.Button("选择输出文件夹", 155);
            chooseOutput.Click += delegate { string path = HostDialogService.Folder(dialog, "选择结果输出文件夹", true); if (!string.IsNullOrEmpty(path)) output.Text = path; };
            AddRow("输入文件夹", input, chooseInput, loadInput); AddRow("使用预设", presetPath, choosePreset, importPreset); AddRow("输出文件夹", output, chooseOutput);
            var chooseStructure=Ui.Button("从架构表读取",155);var resetStructure=Ui.Button("使用文件夹名单",135);
            chooseStructure.Click+=delegate{Attempt(delegate{
                var tasks=CompanyStructurePicker.Choose(dialog,delegate(string path){return CompanyWorkbookHost.FromOpenOrDisk(app,path);},StructureUsesLevel);
                if(tasks==null)return;CompanyStructureTasks.ResolveFiles(tasks,Path.GetTempPath());structureTasks=tasks;files.Rows.Clear();loadedInput=null;
                Message("已选择架构任务"+tasks.Count+"项；按过程表文件名称精确匹配，各文件仍使用自己的Lead。");
                if(Directory.Exists(input.Text.Trim().Trim('"')))LoadFiles();else Message("请输入数据文件夹，然后读取文件。");
            });};
            resetStructure.Click+=delegate{Attempt(delegate{structureTasks=null;files.Rows.Clear();loadedInput=null;LoadFiles();});};
            AddRow("任务名单",new Label{Text="默认选择末级公司，可手动补选其他公司",AutoSize=true},chooseStructure,resetStructure);
            AddRow("路径来源工作表", pathSheet); AddRow("路径来源单元格", pathCell);
            files.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "处理", Width = 60 });
            files.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "文件", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            files.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "状态", ReadOnly = true, Width = 200 }); Ui.StyleGrid(files);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(16) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 50)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 50)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var actions = Ui.ActionBar(); actions.Controls.AddRange(new Control[] { preview, execute, stop });
            root.Controls.Add(settings, 0, 0); root.Controls.Add(files, 0, 1); root.Controls.Add(progress, 0, 2); root.Controls.Add(log, 0, 3); root.Controls.Add(actions, 0, 4); dialog.Controls.Add(root);
            preview.Click += delegate { Attempt(delegate { Begin(false); }); }; execute.Click += delegate { Attempt(delegate { Begin(true); }); };
            stop.Enabled = false; stop.Click += delegate { stopping = true; Message("正在停止；当前临时副本将丢弃，已经输出的成功文件保留。"); };
            timer.Tick += delegate { Advance(); };
            dialog.FormClosing += delegate(object sender, FormClosingEventArgs e) { if (queue != null) { e.Cancel = true; stopping = true; Message("正在停止，请等待当前操作结束后关闭。"); } };
        }
        private void AddRow(string label, Control value, Control action = null, Control secondary = null)
        {
            int row = settings.RowCount++; settings.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            settings.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 7, 12, 7) }, 0, row);
            value.Dock = DockStyle.None; value.Anchor = AnchorStyles.Left | AnchorStyles.Right; value.Margin = new Padding(0, 4, 12, 4); settings.Controls.Add(value, 1, row);
            if (action != null) { action.Anchor = AnchorStyles.Left; action.Margin = new Padding(0, 3, 8, 3); settings.Controls.Add(action, 2, row); }
            if (secondary != null) { secondary.Anchor = AnchorStyles.Left; secondary.Margin = new Padding(0, 3, 0, 3); settings.Controls.Add(secondary, 3, row); }
        }
        private void LoadFiles()
        {
            if (queue != null) return;
            string path = CellPath(input.Text);
            var found = structureTasks==null?Directory.GetFiles(path).Where(IsWorkbook).Where(p => !Path.GetFileName(p).StartsWith("~$", StringComparison.Ordinal)).OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray():CompanyStructureTasks.ResolveFiles(structureTasks,path).ToArray();
            input.Text = path; files.Rows.Clear();
            foreach (string file in found) { int row = files.Rows.Add(structureTasks!=null||DefaultSelected(file), Path.GetFileName(file), File.Exists(file)?"待预检":"未找到文件"); files.Rows[row].Tag = file; }
            loadedInput = path;
            Message(structureTasks==null?"已列出 " + found.Length + " 个工作簿，默认勾选文件名含（合并）或(合并)的合并主体；可手动调整。":"已按架构表顺序匹配"+found.Length+"个任务；未找到的文件会明确报告，不使用近似名称替代。");
        }
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
        internal void UsePreset(BatchPreset value, string prefix)
        {
            if (queue != null) throw new InvalidOperationException("执行期间不能更换预设。");
            var snapshot = BatchPresetStore.Decode(BatchPresetStore.Encode(value), "summary");
            SummaryAreaSettings.Read(snapshot);
            chosenPreset = snapshot; presetPath.Text = prefix + snapshot.Name;
            foreach (DataGridViewRow row in files.Rows) row.Cells[2].Value = "待预检";
            Message("已选择预设“" + snapshot.Name + "”；仅使用坐标配置，各文件读取自己的数据，无需原预设工作簿存在。");
        }
        private void Message(string text) { log.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + text + Environment.NewLine); log.SelectionStart = log.TextLength; log.ScrollToCaret(); }
        private void Attempt(Action action) { try { action(); } catch (Exception error) { Message("失败：" + error.Message); } }
        private void Begin(bool save)
        {
            if (queue != null) return; files.EndEdit();
            if (loadedInput == null) throw new InvalidOperationException("输入路径后，请先按 Enter 或点击读取文件，再勾选工作簿。");
            var selected = files.Rows.Cast<DataGridViewRow>().Where(r => object.Equals(r.Cells[0].Value, true)).ToList();
            if (selected.Count == 0) throw new InvalidOperationException("请勾选需要处理的文件。");
            if (chosenPreset == null) throw new InvalidOperationException("请选择已保存预设，或导入一个预设文件。");
            var selectedPreset = BatchPresetStore.Decode(BatchPresetStore.Encode(chosenPreset), "summary"); SummaryAreaSettings.Read(selectedPreset);
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
            scratch = Path.Combine(destination, ".xybatch-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(scratch);
            workingFile = Path.Combine(scratch, ".xybatch-" + Guid.NewGuid().ToString("N") + Path.GetExtension(inputFile));
            bool copiedOpen = false;
            using (var bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)app).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++)
                {
                    dynamic existing = bag.Own(books.Item(i));
                    if (!string.Equals(Convert.ToString(existing.FullName), Path.GetFullPath(inputFile), StringComparison.OrdinalIgnoreCase)) continue;
                    using (var quiet = new CompanyWorkbookQuiet(app)) existing.SaveCopyAs(workingFile);
                    copiedOpen = true;
                    Message("读取已打开工作簿的当前内容（包含未保存修改）：" + inputFile + "；不保存、不关闭原工作簿。");
                    break;
                }
            }
            if (!copiedOpen) using (var source = new FileStream(inputFile, FileMode.Open, FileAccess.Read, FileShare.Read))
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
                        VerifyNoTemporaryLinks();
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
        private void VerifyNoTemporaryLinks()
        {
            object links = ((dynamic)workBook).LinkSources(1);
            var items = links as Array;
            if (items != null) foreach (object link in items)
                if (Convert.ToString(link).IndexOf(Path.GetFileName(workingFile), StringComparison.OrdinalIgnoreCase) >= 0 || Convert.ToString(link).IndexOf(scratch, StringComparison.OrdinalIgnoreCase) >= 0)
                    throw new InvalidOperationException("输出包含指向临时副本的外链，已阻止保存，请检查模板公式。");
        }
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
