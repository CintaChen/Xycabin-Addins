using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class DataMigrationForm : Form, IDialogDefaultLayout
    {
        internal readonly TextBox SourcePath = new TextBox { Width = 160, Name = "migrationSource" };
        internal readonly ComboBox SourceSheet = new ComboBox { Width = 80, DropDownWidth = 180, Name = "migrationSheet", DropDownStyle = ComboBoxStyle.DropDown };
        internal readonly TextBox SourceRange = new TextBox { Width = 80, Name = "migrationRange" };
        private readonly CheckBox allowEmpty = new CheckBox { Text = "允许全空白源区域（清空目标）", AutoSize = true };
        private MigrationRequest lastRequest;
        internal readonly ResultGrid Targets = new ResultGrid { Name = "migrationTargets", Dock = DockStyle.Fill };
        private readonly RichTextBox details = new RichTextBox { Name = "migrationDetails", ReadOnly = true, Dock = DockStyle.Fill, BackColor = Color.White, WordWrap = true };
        private readonly Label status = Ui.Label("请选择源文件，或点击“读取当前选区”；选择目标后先预检。", 9, false);
        private readonly RoundedButton preview = Ui.Button("预检与预览", 120), execute = Ui.Button("执行迁移", 110), cancel = Ui.Button("取消操作", 110), open = Ui.Button("打开结果文件夹", 150);
        private readonly List<Control> inputs = new List<Control>();
        private readonly Timer pump = new Timer { Interval = 15 };
        private IEnumerator<MigrationProgress> steps;
        private Action finished;
        private Action<Exception> failed;
        private bool syncing, closeWhenDone, ticking, fittingTargets;
        private static readonly int[] TargetColumnWeights = { 32, 13, 9, 11, 9, 18, 20 };
        internal bool Running { get { return steps != null; } }
        internal event Action PreviewRequested, ExecuteRequested, CancelRequested, SelectionRequested, TargetSelectionRequested, InputsChanged, SourceRangeRequested;
        internal event Action<string> SourcePicked;
        internal event Action<string[]> TargetsPicked;

        internal DataMigrationForm()
        {
            Text = "批量迁移数据"; DialogSizing.Prepare(this, "migration.data", "table", Ui.Blue);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Padding = new Padding(16) };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 58)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 42)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var source = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top, WrapContents = true, Margin = Padding.Empty };
            source.Controls.Add(SourceField(SourcePath, "源工作簿")); source.Controls.Add(SourceField(SourceSheet, "源工作表")); source.Controls.Add(SourceField(SourceRange, "源区域"));
            var pickRange = Ui.Button("在表中选择…", 104); pickRange.Name = "migrationPickRange";
            pickRange.Click += delegate { if (SourceRangeRequested != null) SourceRangeRequested(); };
            source.Controls.Add(pickRange);
            var pick = Ui.Button("选择源文件…", 104); pick.Click += delegate { PickSource(); };
            var selection = Ui.Button("读取当前选区", 116); selection.Click += delegate { if (SelectionRequested != null) SelectionRequested(); };
            source.Controls.Add(pick); source.Controls.Add(selection); source.Controls.Add(allowEmpty); CenterSourceRow(source); root.Controls.Add(source, 0, 0);
            var targetActions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top, WrapContents = true };
            var add = Ui.Button("添加目标…", 110); add.Click += delegate { PickTargets(); };
            var remove = Ui.Button("移除目标", 110); remove.Click += delegate { foreach (DataGridViewRow row in Targets.SelectedRows) if (!row.IsNewRow) Targets.Rows.Remove(row); Changed(); };
            var targetSelection = Ui.Button("当前选区设为目标", 165); targetSelection.Click += delegate { if (TargetSelectionRequested != null) TargetSelectionRequested(); };
            var all = Ui.Button("全选", 72); var none = Ui.Button("全不选", 84);
            all.Click += delegate { SelectTargets(true); }; none.Click += delegate { SelectTargets(false); };
            targetActions.Controls.Add(add); targetActions.Controls.Add(remove); targetActions.Controls.Add(targetSelection); targetActions.Controls.Add(all); targetActions.Controls.Add(none); root.Controls.Add(targetActions, 0, 1);
            Targets.AllowUserToAddRows = false; Targets.AllowUserToDeleteRows = false; Targets.RowHeadersVisible = false; Targets.SelectionMode = DataGridViewSelectionMode.FullRowSelect; Targets.MultiSelect = false;
            AddColumn("path", "目标工作簿", 240, true);
            Targets.Columns.Add(new DataGridViewComboBoxColumn {
                Name = "sheet", HeaderText = "工作表", Width = 150, SortMode = DataGridViewColumnSortMode.NotSortable,
                FlatStyle = FlatStyle.Standard, DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox, DisplayStyleForCurrentCellOnly = false,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleLeft, Padding = Padding.Empty }
            });
            AddColumn("start", "起始单元格", 115, false); AddColumn("range", "目标范围", 150, true); AddColumn("overwrite", "覆盖数量", 110, true); AddColumn("status", "状态与原因", 220, true); AddColumn("output", "结果文件", 260, true);
            // Keep existing data column indices stable; show the new checkbox first.
            Targets.Columns.Add(new DataGridViewCheckBoxColumn { Name = "include", HeaderText = "选择", Width = 56, SortMode = DataGridViewColumnSortMode.NotSortable, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } });
            Targets.Columns["include"].DisplayIndex = 0;
            // Allocate the viewport, rather than sizing to every path or every
            // choice in a worksheet dropdown.
            Targets.ColumnFitOverride = FitTargetColumns;
            Targets.ClientSizeChanged += delegate { Ui.AutoFitColumns(Targets); };
            Targets.RowsRemoved += delegate { Ui.AutoFitColumns(Targets); };
            foreach (var scroll in Targets.Controls.OfType<VScrollBar>())
            { scroll.VisibleChanged += delegate { Ui.AutoFitColumns(Targets); }; scroll.SizeChanged += delegate { Ui.AutoFitColumns(Targets); }; }
            Targets.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            Targets.CurrentCellDirtyStateChanged += delegate { if (Targets.IsCurrentCellDirty) Targets.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            Targets.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e) { if (!syncing && (e.ColumnIndex == 1 || e.ColumnIndex == 2 || e.ColumnIndex == 7)) { if (e.ColumnIndex == 1 && e.RowIndex >= 0) Targets.Rows[e.RowIndex].Cells[1].ToolTipText = "点击表名选择工作表：" + Convert.ToString(Targets.Rows[e.RowIndex].Cells[1].Value); Changed(); RefreshResults(); } };
            Targets.SelectionChanged += delegate { ShowDetails(); };
            Targets.EditingControlShowing += delegate(object sender, DataGridViewEditingControlShowingEventArgs e)
            {
                var combo = e.Control as DataGridViewComboBoxEditingControl;
                if (combo == null || Targets.CurrentCell == null || Targets.CurrentCell.ColumnIndex != 1) return;
                combo.FlatStyle = FlatStyle.Standard;
                combo.Font = Targets.CurrentCell.InheritedStyle.Font;
                combo.ForeColor = Targets.CurrentCell.InheritedStyle.ForeColor;
                combo.BackColor = Color.White;
                // A native non-drawn ComboBox fixes its height to the font and
                // shrinks within a taller row. Keep the same font and give the
                // selection field the cell's height through fixed item sizing.
                combo.DrawMode = DrawMode.OwnerDrawFixed;
                int height = Math.Max(combo.Font.Height, Targets.CurrentCell.Size.Height - 1);
                combo.ItemHeight = Math.Max(combo.Font.Height, combo.ItemHeight + height - combo.PreferredHeight);
                combo.DrawItem -= DrawSheetItem; combo.DrawItem += DrawSheetItem;
                int width = Targets.CurrentCell.Size.Width;
                foreach (object item in combo.Items) width = Math.Max(width, TextRenderer.MeasureText(Convert.ToString(item), combo.Font).Width + Ui.Px(Targets, 40));
                combo.DropDownWidth = Math.Min(width, Math.Max(Ui.Px(Targets, 104), Screen.FromControl(Targets).WorkingArea.Width - Ui.Px(Targets, 40)));
                ((DataGridViewComboBoxCell)Targets.CurrentCell).DropDownWidth = combo.DropDownWidth;
                combo.MaxDropDownItems = 12;
            };
            Targets.CellClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (Running || Targets.ReadOnly || e.RowIndex < 0 || e.ColumnIndex != 1) return;
                Targets.CurrentCell = Targets.Rows[e.RowIndex].Cells[1];
                Targets.BeginEdit(true);
                var combo = Targets.EditingControl as DataGridViewComboBoxEditingControl;
                if (combo != null && combo.Items.Count > 0) combo.DroppedDown = true;
            };
            Ui.StyleGrid(Targets);
            Action applyCellPadding = delegate
            {
                // Padding.Empty on a column inherits the grid's padding. Put
                // spacing on the other columns instead so the sheet editor
                // uses the full cell width, just like the closed selector.
                Targets.DefaultCellStyle.Padding = Padding.Empty;
                int padding = Ui.Px(Targets, 8);
                foreach (DataGridViewColumn column in Targets.Columns)
                    if (column.Name != "sheet") column.DefaultCellStyle.Padding = new Padding(padding, 0, padding, 0);
            };
            Targets.HandleCreated += delegate { applyCellPadding(); };
            if (Targets.IsHandleCreated) applyCellPadding();
            root.Controls.Add(Targets, 0, 2); root.Controls.Add(details, 0, 3);
            status.AutoSize = true; status.MaximumSize = new Size(1100, 0); root.Controls.Add(status, 0, 4);
            var actions = Ui.ActionBar(); var close = Ui.Button("关闭", 90);
            preview.Click += delegate { if (PreviewRequested != null) PreviewRequested(); };
            execute.Click += delegate { if (ExecuteRequested != null) ExecuteRequested(); };
            cancel.Click += delegate { if (CancelRequested != null) CancelRequested(); status.Text = "正在取消，等待当前步骤收尾…"; cancel.Enabled = false; };
            open.Click += delegate { OpenResult(); }; close.Click += delegate { Close(); };
            actions.Controls.Add(preview); actions.Controls.Add(execute); actions.Controls.Add(cancel); actions.Controls.Add(open); actions.Controls.Add(close); root.Controls.Add(actions, 0, 5); Controls.Add(root);
            inputs.AddRange(new Control[] { SourcePath, SourceSheet, SourceRange, allowEmpty, pick, pickRange, selection, add, remove, targetSelection, all, none });
            allowEmpty.CheckedChanged += delegate { Changed(); };
            SourcePath.TextChanged += delegate { Changed(); }; SourceSheet.TextChanged += delegate { Changed(); }; SourceRange.TextChanged += delegate { Changed(); };
            pump.Tick += delegate { Tick(); }; execute.Enabled = false; cancel.Enabled = false; open.Enabled = false;
            FormClosing += delegate(object sender, FormClosingEventArgs e) { if (Running && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; closeWhenDone = true; if (CancelRequested != null) CancelRequested(); } };
            Disposed += delegate { pump.Stop(); if (steps != null) { steps.Dispose(); steps = null; } pump.Dispose(); };
            if (!DialogPreviewFactory.IsConstructing) Load += delegate { CompactLegacySourceWidths(); };
        }
        private void CompactLegacySourceWidths()
        {
            // Adjust only the former default widths in this window's copy of
            // the appearance. Other custom widths and the saved file survive.
            var data = FormatSession.Capture(this).Copy(); bool changed = false;
            var legacy = new Dictionary<Control, int> { { SourcePath, 420 }, { SourceSheet, 160 }, { SourceRange, 150 } };
            foreach (var target in DialogFormatting.Register(this, "migration.data"))
            {
                int width; FormatRule rule;
                if (legacy.TryGetValue(target.Control, out width) && data.Formats.TryGetValue("migration.data::" + target.Id, out rule) && rule.Width96 == width)
                { rule.Width96 = 0; changed = true; }
            }
            if (changed) DialogFormatting.Apply(this, "migration.data", data);
        }
        private static Panel SourceField(Control input, string caption)
        {
            var field = new Panel { Margin = new Padding(0, 0, 8, 0) };
            var label = new Label { Text = caption, AutoSize = true };
            field.Controls.Add(label); field.Controls.Add(input);
            var grip = new WidthGrip(input);
            bool arranging = false;
            field.Layout += delegate
            {
                if (arranging) return;
                arranging = true;
                try
                {
                    int height = Math.Max(label.Height, input.Height);
                    label.Location = new Point(3, (height - label.Height) / 2);
                    input.Location = new Point(label.Right + 9, (height - input.Height) / 2);
                    field.Size = new Size(input.Right + grip.Width + 3, height);
                }
                finally { arranging = false; }
            };
            input.SizeChanged += delegate { field.PerformLayout(); };
            label.SizeChanged += delegate { field.PerformLayout(); };
            field.PerformLayout(); return field;
        }
        private static void CenterSourceRow(FlowLayoutPanel row)
        {
            bool aligning = false;
            row.Layout += delegate
            {
                if (aligning || row.Controls.Count == 0) return;
                aligning = true;
                try
                {
                    int height = row.Controls.Cast<Control>().Max(c => c.Height);
                    foreach (Control control in row.Controls)
                    {
                        int gap = height - control.Height;
                        var margin = new Padding(control.Margin.Left, gap / 2, control.Margin.Right, gap - gap / 2);
                        if (control.Margin != margin) control.Margin = margin;
                    }
                }
                finally { aligning = false; }
            };
        }
        internal const string FileFilter = "工作簿 (*.xlsx;*.xlsm;*.xls;*.et)|*.xlsx;*.xlsm;*.xls;*.et";
        private static void DrawSheetItem(object sender, DrawItemEventArgs e)
        {
            var combo = (ComboBox)sender; e.DrawBackground();
            string text = e.Index >= 0 && e.Index < combo.Items.Count ? combo.GetItemText(combo.Items[e.Index]) : combo.Text;
            var bounds = e.Bounds; int inset = Ui.Px(combo, 3); bounds.X += inset; bounds.Width = Math.Max(0, bounds.Width - inset * 2);
            TextRenderer.DrawText(e.Graphics, text, combo.Font, bounds, e.ForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
            e.DrawFocusRectangle();
        }
        private void FitTargetColumns()
        {
            if (fittingTargets || Targets.ClientSize.Width < 1) return;
            fittingTargets = true;
            try
            {
                Ui.EnsureGridHeaders(Targets);
                Targets.Columns["sheet"].MinimumWidth = Math.Max(Targets.Columns["sheet"].MinimumWidth, Ui.Px(Targets, 104));
                // Header minima are pinned first; distribute remaining pixels
                // proportionally. This avoids Fill's width drift at minima.
                for (int pass = 0; pass < 3; pass++)
                {
                    int available = TargetViewportWidth();
                    int[] widths = Targets.Columns.Cast<DataGridViewColumn>().Select(c => c.MinimumWidth).ToArray();
                    var pending = Enumerable.Range(0, TargetColumnWeights.Length).Where(i => Targets.Columns[i].Visible).ToList();
                    int remaining = available - (Targets.Columns["include"].Visible ? widths[7] : 0);
                    int weight = pending.Sum(i => TargetColumnWeights[i]);
                    while (pending.Count > 0)
                    {
                        bool pinned = false;
                        foreach (int index in pending.ToArray())
                            if ((double)remaining * TargetColumnWeights[index] / weight < widths[index])
                            { remaining -= widths[index]; weight -= TargetColumnWeights[index]; pending.Remove(index); pinned = true; }
                        if (pinned) continue;
                        int assigned = 0;
                        foreach (int index in pending)
                        {
                            widths[index] = index == pending[pending.Count - 1] ? remaining - assigned : (int)((double)remaining * TargetColumnWeights[index] / weight);
                            assigned += widths[index];
                        }
                        break;
                    }
                    foreach (DataGridViewColumn column in Targets.Columns)
                        if (column.Visible && column.Width != widths[column.Index]) column.Width = widths[column.Index];
                    if (TargetViewportWidth() == available) break;
                }
            }
            finally { fittingTargets = false; }
        }
        private int TargetViewportWidth()
        {
            return Math.Max(0, Targets.ClientSize.Width - (Targets.RowHeadersVisible ? Targets.RowHeadersWidth : 0)
                - Targets.Controls.OfType<VScrollBar>().Where(c => c.Visible).Sum(c => c.Width));
        }
        private void AddColumn(string name, string title, int width, bool readOnly)
        { Targets.Columns.Add(new DataGridViewTextBoxColumn { Name = name, HeaderText = title, Width = width, ReadOnly = readOnly, SortMode = DataGridViewColumnSortMode.NotSortable }); }
        private void PickSource()
        { string[] paths = HostDialogService.OpenFiles(this, FileFilter); if (paths.Length > 0 && SourcePicked != null) SourcePicked(paths[0]); }
        private void PickTargets()
        {
            try
            {
                string[] selections = MigrationTargetPicker.Show(this); if (selections.Length == 0) return;
                var scan = MigrationTargetScan.Expand(selections, SourcePath.Text.Trim());
                if (scan.Files.Count > 0 && TargetsPicked != null) TargetsPicked(scan.Files.ToArray());
                Message("识别到 " + scan.Files.Count + " 个工作簿，已勾选 " + Targets.Rows.Cast<DataGridViewRow>().Count(IsIncluded) + "/" + Targets.Rows.Count + " 个目标。文件夹只扫描当前层；已跳过 " + scan.Skipped + " 个源文件、临时或迁移结果文件。" + string.Join("；", scan.Errors.ToArray()));
            }
            catch (Exception error) { Message("添加目标失败：" + error.Message); }
        }
        internal static bool IsIncluded(DataGridViewRow row) { return object.Equals(row.Cells["include"].Value, true); }
        internal void SelectTargets(bool selected)
        {
            if (Running) return;
            Targets.EndEdit(); syncing = true;
            try { foreach (DataGridViewRow row in Targets.Rows) row.Cells["include"].Value = selected; }
            finally { syncing = false; }
            Changed(); RefreshResults();
        }
        private void Changed()
        { if (syncing || Running) return; lastRequest = null; execute.Enabled = false; status.Text = "已勾选 " + Targets.Rows.Cast<DataGridViewRow>().Count(IsIncluded) + "/" + Targets.Rows.Count + " 个目标，请明确源区域和目标位置，再重新预检。"; if (InputsChanged != null) InputsChanged(); }
        internal void SetSource(string path, string sheet, string address, IEnumerable<string> sheets)
        {
            syncing = true;
            try { SourcePath.Text = path; SourceSheet.Items.Clear(); foreach (string name in sheets) SourceSheet.Items.Add(name); SourceSheet.Text = sheet; SourceRange.Text = address; }
            finally { syncing = false; } Changed();
        }
        internal void AddTarget(string path, IEnumerable<string> sheets, string selected, string start, string error = "")
        {
            path = Path.GetFullPath(path);
            foreach (DataGridViewRow old in Targets.Rows) if (string.Equals(Convert.ToString(old.Cells[0].Value), path, StringComparison.OrdinalIgnoreCase)) { Targets.CurrentCell = old.Cells[1]; return; }
            syncing = true;
            try
            {
                int index = Targets.Rows.Add(path, null, start, "", "", "待预检", "", error.Length == 0); var row = Targets.Rows[index];
                var combo = (DataGridViewComboBoxCell)row.Cells[1]; foreach (string name in sheets) combo.Items.Add(name);
                if (combo.Items.Contains(selected)) combo.Value = selected;
                else { selected = ""; row.Cells[5].Value = "请选择目标工作表"; }
                row.Tag = new MigrationTarget { Path = path, Sheet = selected, Start = start, Status = error.Length == 0 ? "待预检" : "读取失败", Error = error };
                row.Cells[0].ToolTipText = path;
                row.Cells[1].ToolTipText = "点击表名选择工作表：" + selected;
                if (error.Length > 0) { row.Cells[5].Value = "未勾选：" + error; row.Cells[5].ToolTipText = error; }
            }
            finally { syncing = false; } Changed(); Ui.AutoFitColumns(Targets);
        }
        internal MigrationRequest Request()
        {
            Targets.EndEdit();
            if (SourceRange.Text.Trim().Length == 0) throw new InvalidOperationException("请填写源区域（如 B5:I22），或点击“读取当前选区”；不会默认迁移 A1。");
            MigrationArea.Parse(SourceRange.Text.Trim());
            var request = new MigrationRequest { SourcePath = Path.GetFullPath(SourcePath.Text.Trim()), SourceSheet = SourceSheet.Text.Trim(), SourceRange = SourceRange.Text.Trim(), AllowEmptySource = allowEmpty.Checked };
            if (request.SourceSheet.Length == 0 || Targets.Rows.Count == 0) throw new InvalidOperationException("请选择源工作表并添加至少一个目标。");
            foreach (DataGridViewRow row in Targets.Rows)
            {
                if (!IsIncluded(row)) continue;
                var target = new MigrationTarget { Path = Convert.ToString(row.Cells[0].Value), Sheet = Convert.ToString(row.Cells[1].Value), Start = Convert.ToString(row.Cells[2].Value) };
                row.Tag = target; request.Targets.Add(target);
            }
            if (request.Targets.Count == 0) throw new InvalidOperationException("请至少勾选一个目标工作簿。");
            lastRequest = request; return request;
        }
        internal void RefreshResults()
        {
            syncing = true;
            try
            {
                bool ready = false, hasOutput = false;
                foreach (DataGridViewRow row in Targets.Rows)
                {
                    var target = row.Tag as MigrationTarget; if (target == null) continue;
                    row.Cells[3].Value = target.Area == null ? "" : target.Area.Address; row.Cells[4].Value = target.Overwritten;
                    row.Cells[5].Value = (IsIncluded(row) ? "" : "未勾选；") + target.Status + (target.Error.Length == 0 ? "" : "：" + target.Error); row.Cells[5].ToolTipText = target.Error;
                    row.Cells[6].Value = target.Output; row.Cells[6].ToolTipText = target.Output; ready |= IsIncluded(row) && lastRequest != null && lastRequest.Targets.Contains(target) && target.Ready && target.Status == "可执行"; hasOutput |= target.Output.Length > 0;
                }
                execute.Enabled = !Running && ready; open.Enabled = !Running && hasOutput;
            }
            finally { syncing = false; } Ui.AutoFitColumns(Targets); ShowDetails();
            if (!Running && lastRequest != null) status.Text = Summary(lastRequest);
        }
        internal static string Summary(MigrationRequest request)
        {
            int succeeded = request.Targets.Count(t => t.Status == "成功"), failed = request.Targets.Count(t => t.Status == "失败"), ready = request.Targets.Count(t => t.Ready && t.Status == "可执行");
            return "源 " + request.SourceSheet + "!" + request.SourceRange + "：非空 " + request.NonEmptyCells + "，公式 " + request.FormulaCells
                + "；目标 " + request.Targets.Count + "：成功 " + succeeded + "，失败 " + failed + "，可执行 " + ready + "。每个成功目标对应一个结果副本。";
        }
        private void ShowDetails()
        {
            if (Targets.CurrentRow == null) return; var target = Targets.CurrentRow.Tag as MigrationTarget; if (target == null) return;
            details.Text = target.Path + Environment.NewLine + (IsIncluded(Targets.CurrentRow) ? "已勾选，参与本轮预检与迁移" : "未勾选，不参与本轮预检与迁移") + Environment.NewLine + "状态：" + target.Status + Environment.NewLine + target.Error + Environment.NewLine
                + "输出：" + target.Output + Environment.NewLine + "必要名称：" + (target.Names == null ? 0 : target.Names.Count) + Environment.NewLine
                + string.Join(Environment.NewLine, target.Samples.ToArray());
        }
        internal void Run(IEnumerable<MigrationProgress> sequence, Action complete, Action<Exception> error)
        {
            if (Running) throw new InvalidOperationException("当前任务尚未结束。");
            finished = complete; failed = error; steps = sequence.GetEnumerator();
            foreach (Control input in inputs) input.Enabled = false; Targets.ReadOnly = true; preview.Enabled = execute.Enabled = open.Enabled = false; cancel.Enabled = true;
            pump.Start();
        }
        private void Tick()
        {
            if (steps == null || ticking) return;
            // Native COM calls can dispatch Windows messages on this same STA.
            // A nested timer tick must never advance/dispose the same iterator.
            ticking = true; pump.Stop();
            try
            {
                var clock = Stopwatch.StartNew();
                do { if (!steps.MoveNext()) { Finish(null); return; } status.Text = steps.Current.Message; } while (clock.ElapsedMilliseconds < 12);
            }
            catch (Exception error) { Finish(error); }
            finally { ticking = false; if (!IsDisposed && !Disposing && steps != null) pump.Start(); }
        }
        private void Finish(Exception error)
        {
            pump.Stop(); var current = steps; steps = null;
            try { if (current != null) current.Dispose(); } catch (Exception cleanup) { error = error ?? cleanup; }
            foreach (Control input in inputs) input.Enabled = true; Targets.ReadOnly = false; preview.Enabled = true; cancel.Enabled = false;
            if (error == null) { status.Text = lastRequest == null ? "请查看各目标结果。" : Summary(lastRequest); if (finished != null) finished(); }
            else { status.Text = error.Message; if (failed != null) failed(error); }
            RefreshResults(); if (error != null) status.Text = error.Message; if (closeWhenDone) Close();
        }
        internal void Message(string text) { status.Text = text; }
        private void OpenResult()
        {
            var target = Targets.CurrentRow == null ? null : Targets.CurrentRow.Tag as MigrationTarget;
            if (target == null || !File.Exists(target.Output)) target = Targets.Rows.Cast<DataGridViewRow>().Select(r => r.Tag as MigrationTarget).FirstOrDefault(t => t != null && File.Exists(t.Output));
            if (target != null) Process.Start(new ProcessStartInfo("explorer.exe", "/select,\"" + target.Output + "\"") { UseShellExecute = true });
        }
        public Size MeasureDefaultClient() { return new Size(1180, 760); }
    }
}
