using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // =====================================================================
    // 跨工作簿表头匹配合并 —— 界面
    //
    // 只做参数收集与结果呈现；一切 COM 通过事件抛回 Connect，在 Ribbon
    // 宿主线程执行（不继承 FileToolForm —— 它的 Run<T> 会开后台 STA 线程，
    // 跨线程调 WPS COM 是项目红线）。
    //
    // 沿用「取消操作 / 确认操作」二段式语义：写入后保留可还原标记，
    // 用户点取消即用内存快照还原；点确认则结束本批（文件备份保留）。
    // =====================================================================

    internal sealed class MergeHeadersForm : Form, IDialogDefaultLayout
    {
        // 路径可手打，也保留「选择目标…」浏览；默认由 Connect 填成当前活动工作簿。
        // 要挂拖拽手柄就不能用 Anchor=Left|Right（拖出来的宽度会被布局覆盖），改固定宽。
        private readonly TextBox targetPath = new TextBox { Width = 420, Name = "targetWorkbook" };
        private readonly ComboBox targetSheet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Width = 190, Name = "targetSheetBox" };
        private readonly TextBox targetHeaderRow = new TextBox { Width = 60, Text = "1", Name = "targetHeaderRowBox" };
        private readonly Label missingLabel = new Label();
        private readonly Label hint = Ui.Label("", 9, true);
        private readonly RichTextBox logView = new RichTextBox
        { ReadOnly = true, WordWrap = true, ScrollBars = RichTextBoxScrollBars.Vertical, Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Color.White };
        private readonly FlowLayoutPanel actions = Ui.ActionBar();
        private readonly TableLayoutPanel root;
        internal readonly ResultGrid Sources = new ResultGrid();
        internal readonly ResultGrid Map = new ResultGrid();
        internal readonly ResultGrid Samples = new ResultGrid();
        internal readonly RoundedButton Preview, Execute, Stop, Confirm, CopyLog, CloseButton;

        private MergeMatchResult match;
        private bool busy;
        private bool canUndo;
        private bool missingAcknowledged;

        internal event Action PreviewRequested;
        internal event Action ExecuteRequested;
        internal event Action UndoRequested;
        internal event Action ConfirmRequested;
        /// <summary>选定目标工作簿后，请求读一次它的工作表名（COM 只能在宿主线程做）。</summary>
        internal event Action<string> TargetSheetsRequested;
        /// <summary>加入一个来源工作簿后，请求读一次它的工作表名。</summary>
        internal event Action<string> SourceSheetsRequested;

        internal MergeHeadersForm()
        {
            Text = "跨工作簿表头匹配合并";
            DialogSizing.Prepare(this, "merge.headers", "table", Ui.Blue);
            KeyPreview = true;

            root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 13, Padding = new Padding(16) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 0  目标区
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));   // 1  手柄：目标区 ↔ 来源
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 2  提示
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 3  缺失列提示
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 150)); // 4  来源
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));   // 5  手柄：来源 ↔ 列映射
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 45));   // 6  列映射
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));   // 7  手柄：列映射 ↔ 公式样本
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 20));   // 8  公式样本
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));   // 9  手柄：公式样本 ↔ 报告
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 35));   // 10 报告
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));   // 11 手柄：报告下边界（按钮栏固定）
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 12 按钮

            // ---- 区域 1：目标工作簿 ----
            // 每个输入框/下拉框/按钮都带右侧拖拽手柄（Ui.Grip 打包「标签 + 控件 + 手柄」），
            // 拖完的宽度经 DialogFormatting 存进本机自定配置，下次打开仍保持。
            TableLayoutPanel targetRow = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4, Name = "targetRow" };
            targetRow.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            RoundedButton pickTarget = Ui.Button("选择目标…", 108);
            pickTarget.Name = "pickTargetButton";
            pickTarget.Click += delegate { PickFile(false); };
            targetRow.Controls.Add(Ui.Grip(targetPath, "目标工作簿"), 0, 0);
            targetRow.Controls.Add(Ui.Grip(pickTarget, (Control)null), 1, 0);
            targetRow.Controls.Add(Ui.Grip(targetSheet, "工作表"), 2, 0);
            targetRow.Controls.Add(Ui.Grip(targetHeaderRow, "表头行"), 3, 0);
            foreach (ColumnStyle style in new ColumnStyle[] {
                new ColumnStyle(SizeType.AutoSize), new ColumnStyle(SizeType.AutoSize),
                new ColumnStyle(SizeType.AutoSize), new ColumnStyle(SizeType.AutoSize) }) targetRow.ColumnStyles.Add(style);

            // ---- 区域 2：提示 + 缺失列 ----
            // 加粗 + 强制单行：Ui.Label 默认已是 AutoSize=false / AutoEllipsis=true（单行省略），
            // 这里把 MaximumSize 从「限宽导致折行」改成「只限高」（宽 0 = 不限）。
            hint.Margin = new Padding(0, 8, 0, 8);
            // 宽度用 4096 而不是 0：MaximumSize 的语义在这里要显式给足，避免宽度被当成 0。
            hint.MaximumSize = new Size(4096, TextRenderer.MeasureText("测", hint.Font, Size.Empty,
                TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding).Height + 2);
            hint.Text = "按表头列名匹配，把多个来源工作簿的列追加到目标工作簿同名列的已有数据下方。目标与来源都通过文件选择;表头行默认第1行，可改。";
            missingLabel.AutoSize = true; missingLabel.ForeColor = Ui.Danger;
            missingLabel.MaximumSize = new Size(900, 0); missingLabel.Margin = new Padding(0, 0, 0, 8);
            missingLabel.Font = Ui.MakeFont(9, FontStyle.Regular, Font);
            missingLabel.Visible = false;

            // ---- 区域 3：来源 ----
            RoundedButton addSource = Ui.Button("添加来源…", 108);
            RoundedButton removeSource = Ui.Button("移除选中", 100);
            addSource.Name = "addSourceButton";
            removeSource.Name = "removeSourceButton";
            addSource.Click += delegate { PickFile(true); };
            removeSource.Click += delegate { RemoveSelectedSource(); };
            FlowLayoutPanel sourceBar = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Name = "sourceBar" };
            sourceBar.Controls.Add(Ui.Grip(addSource, (Control)null));
            sourceBar.Controls.Add(Ui.Grip(removeSource, (Control)null));
            UiCard sourceCard = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(12), Margin = new Padding(0, 0, 0, 12) };
            BuildSourcesGrid();
            sourceCard.Controls.Add(Sources);
            Panel sourcePanel = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
            sourcePanel.Controls.Add(sourceCard);
            sourcePanel.Controls.Add(sourceBar);

            // ---- 区域 4/5：列映射 + 公式样本 ----
            BuildMapGrid();
            BuildSamplesGrid();
            UiCard mapCard = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(12), Margin = new Padding(0, 0, 0, 12) };
            mapCard.Controls.Add(Map);
            UiCard sampleCard = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(12), Margin = new Padding(0, 0, 0, 12) };
            sampleCard.Controls.Add(Samples);

            // ---- 区域 6：报告 ----
            UiCard logCard = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(12), Margin = new Padding(0, 0, 0, 12) };
            logCard.Controls.Add(logView);

            // ---- 按钮栏 ----
            Preview = Ui.Button("读取表头并预检", 140); Ui.StylePrimary(Preview);
            Execute = Ui.Button("执行合并", 108);
            Stop = Ui.Button("取消操作", 104);
            Confirm = Ui.Button("确认操作", 104);
            CopyLog = Ui.Button("复制日志", 100);
            CloseButton = Ui.Button("关闭", 88);
            actions.Controls.AddRange(new Control[] { Preview, Execute, Stop, Confirm, CopyLog, CloseButton });
            ToolTip tip = new ToolTip();
            tip.SetToolTip(Preview, "读取目标与来源的表头并生成匹配结果（Ctrl+Shift+P）");
            tip.SetToolTip(Execute, "按匹配结果写入目标工作簿（Ctrl+Shift+W）");
            tip.SetToolTip(Stop, "用写入前的快照还原目标工作表");
            tip.SetToolTip(Confirm, "确认本次结果，结束可还原状态（文件备份保留）");

            Preview.Click += delegate { if (PreviewRequested != null) PreviewRequested(); };
            Execute.Click += delegate { if (ExecuteRequested != null) ExecuteRequested(); };
            Stop.Click += delegate { if (busy) { UiDialog.Show(this, "正在处理，请等待当前步骤结束。", Text, MessageBoxButtons.OK, MessageBoxIcon.Information); return; } if (UndoRequested != null) UndoRequested(); };
            Confirm.Click += delegate
            {
                if (UiDialog.Show(this, "确认后本批不再提供还原；程序生成的备份文件保留在目标工作簿同目录，可自行取用或删除。",
                    "确认操作", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.OK) return;
                canUndo = false;
                if (ConfirmRequested != null) ConfirmRequested();
                AppendLog("已确认；本批不再提供还原。");
                RefreshActions();
            };
            CopyLog.Click += delegate
            { try { Clipboard.SetDataObject(logView.Text, true); } catch (Exception ex) { AppendLog("复制失败：" + ex.Message); } };
            CloseButton.Click += delegate { Close(); };

            // 四个方框各自下方一条 8px 拖拽条：拖动时改「上方框」与「下一个内容框」的行高。
            RowGrip gripTarget = new RowGrip(root, 0, 4, "row.target", 92, 120);
            RowGrip gripSources = new RowGrip(root, 4, 6, "row.sources", 120, 120);
            RowGrip gripMap = new RowGrip(root, 6, 8, "row.map", 120, 96);
            RowGrip gripSamples = new RowGrip(root, 8, 10, "row.samples", 96, 120);
            // 报告框（最下方）下边界也支持拖拽：below 指向按钮栏并固定，只改报告框自身高度。
            RowGrip gripReport = new RowGrip(root, 10, 12, "row.report", 120, 0, true);

            UiCard targetCard = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(12), Margin = new Padding(0, 0, 0, 12) };
            targetCard.Controls.Add(targetRow);

            root.Controls.Add(targetCard, 0, 0);
            root.Controls.Add(gripTarget, 0, 1);
            root.Controls.Add(hint, 0, 2);
            root.Controls.Add(missingLabel, 0, 3);
            root.Controls.Add(sourcePanel, 0, 4);
            root.Controls.Add(gripSources, 0, 5);
            root.Controls.Add(mapCard, 0, 6);
            root.Controls.Add(gripMap, 0, 7);
            root.Controls.Add(sampleCard, 0, 8);
            root.Controls.Add(gripSamples, 0, 9);
            root.Controls.Add(logCard, 0, 10);
            root.Controls.Add(gripReport, 0, 11);
            root.Controls.Add(actions, 0, 12);
            Controls.Add(root);

            RefreshActions();
        }

        // ------------------------ 网格 ------------------------

        private void BuildSourcesGrid()
        {
            Sources.Dock = DockStyle.Fill;
            Sources.AllowUserToAddRows = false;
            Sources.AllowUserToDeleteRows = false;
            Sources.AllowUserToResizeRows = false;
            Sources.RowHeadersVisible = false;
            Sources.MultiSelect = false;
            Sources.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            Sources.AutoGenerateColumns = false;
            Sources.ScrollBars = ScrollBars.Both;
            Sources.RowTemplate.Height = 30;
            Sources.EnableHeadersVisualStyles = false;
            Sources.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            Sources.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            Sources.ShowCellToolTips = true;
            DataGridViewComboBoxColumn sheet = new DataGridViewComboBoxColumn { Name = "sheet", HeaderText = "工作表", Width = 160, SortMode = DataGridViewColumnSortMode.NotSortable };
            Sources.Columns.Add(new DataGridViewTextBoxColumn { Name = "workbook", HeaderText = "来源工作簿", ReadOnly = true, Width = 210, SortMode = DataGridViewColumnSortMode.NotSortable });
            Sources.Columns.Add(sheet);
            Sources.Columns.Add(new DataGridViewTextBoxColumn { Name = "headerRow", HeaderText = "表头行", Width = 80, SortMode = DataGridViewColumnSortMode.NotSortable });
            Sources.Columns.Add(new DataGridViewTextBoxColumn { Name = "status", HeaderText = "状态", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 200, SortMode = DataGridViewColumnSortMode.NotSortable });
            // 「工作表」是下拉列且 Items 可能为空：预览/格式渲染路径会灌入与列类型不匹配的示例值，
            // 不吞掉 DataError 会弹出模态错误框等待点击，导致离线探针挂起。
            Sources.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            Ui.StyleGrid(Sources);
        }

        private void BuildMapGrid()
        {
            Map.Dock = DockStyle.Fill;
            Map.AllowUserToAddRows = false;
            Map.AllowUserToDeleteRows = false;
            Map.AllowUserToResizeRows = false;
            Map.RowHeadersVisible = false;
            Map.MultiSelect = false;
            Map.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            Map.AutoGenerateColumns = false;
            Map.ScrollBars = ScrollBars.Both;
            Map.RowTemplate.Height = 28;
            Map.EnableHeadersVisualStyles = false;
            Map.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            Map.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            Map.ShowCellToolTips = true;
            MakeColumn(Map, "source", "来源工作簿", 170, 300);
            MakeColumn(Map, "sheet", "工作表", 120, 200);
            MakeColumn(Map, "header", "表头列名", 150, 260);
            MakeColumn(Map, "srcCol", "源列", 62, 100);
            MakeColumn(Map, "tgtCol", "目标列", 62, 100);
            MakeColumn(Map, "mapping", "映射", 100, 150);
            MakeColumn(Map, "rows", "行数", 66, 100);
            MakeColumn(Map, "startRow", "写入起始行", 96, 140);
            MakeColumn(Map, "status", "状态", 150, 0);
            Map.Columns[Map.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Map.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            Ui.StyleGrid(Map);
        }

        private void BuildSamplesGrid()
        {
            Samples.Dock = DockStyle.Fill;
            Samples.AllowUserToAddRows = false;
            Samples.AllowUserToDeleteRows = false;
            Samples.RowHeadersVisible = false;
            Samples.MultiSelect = false;
            Samples.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            Samples.AutoGenerateColumns = false;
            Samples.ScrollBars = ScrollBars.Both;
            Samples.RowTemplate.Height = 28;
            Samples.EnableHeadersVisualStyles = false;
            Samples.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            Samples.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            Samples.ShowCellToolTips = true;
            MakeColumn(Samples, "source", "来源", 150, 260);
            MakeColumn(Samples, "cell", "源单元格", 96, 140);
            MakeColumn(Samples, "before", "原公式", 240, 420);
            MakeColumn(Samples, "after", "新公式", 240, 420);
            MakeColumn(Samples, "note", "备注", 160, 0);
            Samples.Columns[Samples.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Samples.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            Ui.StyleGrid(Samples);
        }

        // 列宽上限存在 Tag（DataGridViewColumn 没有 MaximumWidth），
        // 由 Ui.AutoFitColumns 读取并封顶；最后一列 Fill 由调用方设置。
        private static void MakeColumn(DataGridView grid, string name, string header, int width, int cap)
        {
            DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
            column.Name = name;
            column.HeaderText = header;
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            column.MinimumWidth = 50;
            column.SortMode = DataGridViewColumnSortMode.NotSortable;
            Font headerFont = grid.ColumnHeadersDefaultCellStyle.Font ?? grid.Font;
            int needed = TextRenderer.MeasureText(header, headerFont).Width + 32;
            column.Width = Math.Max(width, needed);
            if (cap > 0) column.Tag = Math.Max(cap, needed);
            grid.Columns.Add(column);
        }

        // ------------------------ 数据 ------------------------

        /// <summary>从界面收集请求（不含 COM）。</summary>
        internal MergeRequest BuildRequest()
        {
            MergeRequest request = new MergeRequest();
            request.Target.Path = targetPath.Text.Trim();
            request.Target.DisplayName = MergeWorkbookSpec.Display(request.Target.Path);
            request.Target.SheetName = targetSheet.Text.Trim();
            int headerRow;
            if (!int.TryParse(targetHeaderRow.Text.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out headerRow)) headerRow = 1;
            request.Target.HeaderRow = headerRow;
            if (request.Target.SheetNames.Count == 0 && targetSheet.Items.Count > 0)
                foreach (object item in targetSheet.Items) request.Target.SheetNames.Add(Convert.ToString(item));

            foreach (DataGridViewRow row in Sources.Rows)
            {
                MergeWorkbookSpec spec = row.Tag as MergeWorkbookSpec;
                if (spec == null) continue;
                string sheetName = Convert.ToString(row.Cells["sheet"].Value);
                spec.SheetName = sheetName == null ? "" : sheetName.Trim();
                int sourceRow;
                if (!int.TryParse(Convert.ToString(row.Cells["headerRow"].Value), NumberStyles.None, CultureInfo.InvariantCulture, out sourceRow)) sourceRow = 1;
                spec.HeaderRow = sourceRow;
                request.Sources.Add(spec);
            }
            return request;
        }

        private void PickFile(bool asSource)
        {
            if (busy || canUndo) return;
            if (DialogPreviewFactory.IsConstructing) return;
            string[] selected = HostDialogService.OpenFiles(this);
            if (selected == null || selected.Length == 0) return;
            foreach (string path in selected)
            {
                if (!asSource)
                {
                    targetPath.Text = path;
                    targetSheet.Items.Clear();
                    targetSheet.Text = "";
                    if (TargetSheetsRequested != null) TargetSheetsRequested(path);
                    continue;
                }
                if (AddSource(path)) continue;
            }
        }

        private bool AddSource(string path)
        {
            foreach (DataGridViewRow existing in Sources.Rows)
            {
                MergeWorkbookSpec current = existing.Tag as MergeWorkbookSpec;
                if (current == null) continue;
                if (string.Equals(current.Path, path, StringComparison.OrdinalIgnoreCase)) return false;
            }
            MergeWorkbookSpec spec = new MergeWorkbookSpec();
            spec.Path = path;
            spec.DisplayName = MergeWorkbookSpec.Display(path);
            spec.HeaderRow = 1;
            // 工作表下拉不能等预检才填：加进列表的当下就抛给 Connect 去读一次工作表名。
            int index = Sources.Rows.Add(spec.DisplayName, "", "1", "正在读取工作表…");
            Sources.Rows[index].Tag = spec;
            if (SourceSheetsRequested != null) SourceSheetsRequested(path);
            return true;
        }

        /// <summary>Connect 读完 COM 后回写某一来源行的「工作表」下拉。必须先灌 Items 再赋 Value。</summary>
        internal void SetSourceSheets(string path, List<string> names, string error)
        {
            if (IsDisposed) return;
            foreach (DataGridViewRow row in Sources.Rows)
            {
                MergeWorkbookSpec spec = row.Tag as MergeWorkbookSpec;
                if (spec == null || !string.Equals(spec.Path, path, StringComparison.OrdinalIgnoreCase)) continue;
                if (!string.IsNullOrEmpty(error))
                {
                    row.Cells["status"].Value = "读取失败：" + error;
                    return;
                }
                spec.SheetNames.Clear();
                if (names != null) spec.SheetNames.AddRange(names);
                DataGridViewComboBoxCell cell = row.Cells["sheet"] as DataGridViewComboBoxCell;
                if (cell != null)
                {
                    cell.Items.Clear();
                    foreach (string name in spec.SheetNames) cell.Items.Add(name);
                    row.Cells["sheet"].Value = spec.SheetNames.Count > 0 ? spec.SheetNames[0] : "";
                }
                row.Cells["status"].Value = spec.SheetNames.Count == 0 ? "没有可见工作表" : "待预检";
                return;
            }
        }

        /// <summary>Connect 在宿主线程取到当前活动工作簿后回填目标区（路径 + 工作表名）。</summary>
        internal void SetTarget(string path, string sheet)
        {
            if (IsDisposed) return;
            targetPath.Text = path == null ? "" : path;
            if (!string.IsNullOrEmpty(sheet)) targetSheet.Text = sheet;
        }

        /// <summary>回写目标工作簿的工作表下拉。</summary>
        internal void SetTargetSheets(List<string> names, string error)
        {
            if (IsDisposed) return;
            if (!string.IsNullOrEmpty(error)) { AppendLog("读取目标工作簿的工作表失败：" + error); return; }
            targetSheet.Items.Clear();
            if (names == null) return;
            foreach (string name in names) targetSheet.Items.Add(name);
            if (targetSheet.Items.Count > 0) targetSheet.Text = Convert.ToString(targetSheet.Items[0]);
        }

        private void RemoveSelectedSource()
        {
            if (busy || canUndo) return;
            if (Sources.SelectedRows.Count == 0) return;
            Sources.Rows.Remove(Sources.SelectedRows[0]);
            match = null;
            missingAcknowledged = false;
            Map.Rows.Clear();
            Samples.Rows.Clear();
            RefreshActions();
        }

        /// <summary>预检后回填工作表下拉与匹配结果。</summary>
        internal void SetMatch(MergeRequest request, MergeMatchResult result)
        {
            match = result;
            missingAcknowledged = !result.HasMissing;

            if (targetSheet.Items.Count == 0 || request.Target.SheetNames.Count > 0)
            {
                targetSheet.Items.Clear();
                foreach (string name in request.Target.SheetNames) targetSheet.Items.Add(name);
            }
            foreach (DataGridViewRow row in Sources.Rows)
            {
                MergeWorkbookSpec spec = row.Tag as MergeWorkbookSpec;
                if (spec == null) continue;
                DataGridViewComboBoxCell cell = row.Cells["sheet"] as DataGridViewComboBoxCell;
                if (cell == null || spec.SheetNames.Count == 0) continue;
                cell.Items.Clear();
                foreach (string name in spec.SheetNames) cell.Items.Add(name);
            }

            Map.Rows.Clear();
            Samples.Rows.Clear();
            if (result == null) { RefreshActions(); return; }

            for (int i = 0; i < result.Sources.Count && i < request.Sources.Count; i++)
            {
                MergeSourcePlan plan = result.Sources[i];
                MergeWorkbookSpec spec = request.Sources[i];
                Sources.Rows[i].Cells["status"].Value = plan.Maps.Count + " 列匹配 / " + plan.Missing.Count + " 列缺失";
                foreach (MergeColumnMap map in plan.Maps)
                {
                    int index = Map.Rows.Add(spec.DisplayName, plan.SheetName, map.Header, map.SourceLetter,
                        map.TargetLetter, map.SourceLetter + " → " + map.TargetLetter,
                        map.RowCount, map.TargetStartRow, plan.StartRowsDiffer ? "起始行与本来源其他列不同" : "待写入");
                    Map.Rows[index].Tag = map;
                }
                foreach (string missing in plan.Missing)
                    Map.Rows.Add(spec.DisplayName, plan.SheetName, missing, "", "", "", 0, 0, "目标缺失，不写入");
            }
            Ui.AutoFitColumns(Map);

            if (result.HasMissing)
            {
                missingLabel.Text = "⚠ " + MergePlanText.MissingText(result).Replace("\r\n", " ").Trim();
                missingLabel.Visible = true;
            }
            else missingLabel.Visible = false;

            StringBuilder text = new StringBuilder();
            foreach (string blocking in result.Blocking) text.AppendLine("阻断：" + blocking);
            foreach (string warning in result.Warnings) text.AppendLine("提示：" + warning);
            if (result.Blocking.Count == 0)
                text.AppendLine("预检完成：" + result.MappedColumns + " 列匹配、"
                    + result.TotalCells + " 个单元格待写入。请点「执行合并」。");
            SetLog(text.ToString());
            RefreshActions();
        }

        internal void SetReport(MergeReport report)
        {
            if (report == null) return;
            Samples.Rows.Clear();
            foreach (MergeRewriteSample sample in report.Samples)
            {
                if (Samples.Rows.Count >= 50) break;
                Samples.Rows.Add(sample.Source, sample.Cell, sample.Before, sample.After, sample.Note);
            }
            Ui.AutoFitColumns(Samples);
            SetLog(MergePlanText.ReportText(report, MergeWorkbookSpec.Display(targetPath.Text.Trim())));
            if (report.BackupPath.Length > 0) canUndo = true;
            RefreshActions();
        }

        internal void SetSamples(List<MergeRewriteSample> samples)
        {
            Samples.Rows.Clear();
            if (samples == null) return;
            foreach (MergeRewriteSample sample in samples)
            {
                if (Samples.Rows.Count >= 50) break;
                Samples.Rows.Add(sample.Source, sample.Cell, sample.Before, sample.After, sample.Note);
            }
            Ui.AutoFitColumns(Samples);
        }

        internal void SetStatus(string text)
        {
            if (IsDisposed || !IsHandleCreated) return;
            BeginInvoke((MethodInvoker)delegate { hint.Text = text; });
        }

        internal void AppendLog(string text)
        {
            if (IsDisposed || !IsHandleCreated) { return; }
            BeginInvoke((MethodInvoker)delegate
            { logView.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + text + "\r\n"); });
        }

        internal void SetLog(string text)
        { logView.Clear(); logView.AppendText(text); }

        internal void SetBusy(bool value)
        { busy = value; RefreshActions(); }

        internal void MarkUndone()
        { canUndo = false; RefreshActions(); }

        internal bool IsBusy { get { return busy; } }
        internal bool CanUndo { get { return canUndo; } }
        internal MergeMatchResult Match { get { return match; } }
        internal bool MissingAcknowledged
        {
            get { return missingAcknowledged; }
            set { missingAcknowledged = value; RefreshActions(); }
        }

        private void RefreshActions()
        {
            bool idle = !busy && !canUndo;
            targetPath.Enabled = idle;
            Sources.ReadOnly = !idle;
            Preview.Enabled = Execute.Enabled = idle;
            Execute.Enabled = idle && match != null && match.Blocking.Count == 0
                && match.MappedColumns > 0 && missingAcknowledged;
            Stop.Enabled = busy || canUndo;
            Confirm.Enabled = !busy && canUndo;
            CopyLog.Enabled = CloseButton.Enabled = true;
        }

        public Size MeasureDefaultClient()
        {
            root.PerformLayout();
            return new Size(DialogSizing.LogicalClient(this).Width, 620);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { e.Handled = true; e.SuppressKeyPress = true; Close(); }
            if (e.Control && e.Shift && e.KeyCode == Keys.P) { e.Handled = true; if (Preview.Enabled && PreviewRequested != null) PreviewRequested(); }
            if (e.Control && e.Shift && e.KeyCode == Keys.W) { e.Handled = true; if (Execute.Enabled && ExecuteRequested != null) ExecuteRequested(); }
            if (e.Control && e.Shift && e.KeyCode == Keys.C) { e.Handled = true; try { Clipboard.SetDataObject(logView.Text, true); } catch { } }
            base.OnKeyDown(e);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (busy)
            {
                e.Cancel = true;
                AppendLog("正在处理，请等待当前步骤结束再关闭。");
                return;
            }
            if (canUndo)
                AppendLog("本批仍可还原；如需保留现场，请关闭前先点「取消操作」或「确认操作」。");
            base.OnFormClosing(e);
        }

        // 非模态显示，与链接助手/按列筛选等其它弹窗完全一致：
        // 不重写 ShowWithoutActivation、不设 ShowInTaskbar/MinimizeBox，
        // 窗口样式对齐正常弹窗，避免 tool window 行为在 WPS 宿主线程抢焦点/失激活。

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // 读回上次拖过的方框高度（存在本机自定配置里，键名是布局自己用的合成键）。
            RestoreRowHeight(0, 4, "row.target");
            RestoreRowHeight(4, 6, "row.sources");
            RestoreRowHeight(6, 8, "row.map");
            RestoreRowHeight(8, 10, "row.samples");
            RestoreRowHeight(10, 12, "row.report", true);
        }

        private void RestoreRowHeight(int above, int below, string key)
        { RestoreRowHeight(above, below, key, false); }

        private void RestoreRowHeight(int above, int below, string key, bool fixedBelow)
        {
            int saved = DialogFormatting.LoadRowHeight(this, key);
            if (saved <= 0) return;
            // The store contains 96-DPI logical pixels, just like RowGrip.Commit.
            saved = Ui.Px(this, saved);
            int[] heights = root.GetRowHeights();
            if (above >= heights.Length || below >= heights.Length) return;
            if (fixedBelow)
            {
                // 报告框：below 是按钮栏（AutoSize），只恢复 above 高度，不动按钮栏。
                if (saved < Ui.Px(this, 80)) return;
                root.RowStyles[above].SizeType = SizeType.Absolute; root.RowStyles[above].Height = saved;
                root.PerformLayout();
                return;
            }
            int belowHeight = heights[below] + (heights[above] - saved);
            if (belowHeight < Ui.Px(this, 80)) return;   // 空间不够就不恢复，免得把下面的方框压没
            root.RowStyles[above].SizeType = SizeType.Absolute; root.RowStyles[above].Height = saved;
            root.RowStyles[below].SizeType = SizeType.Absolute; root.RowStyles[below].Height = belowHeight;
            root.PerformLayout();
        }
    }
}
