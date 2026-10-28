using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace Excel_Iterate
{
    // =====================================================================
    // 合并抵消（往来抵消 / 收入成本抵消）—— 界面
    //
    // 纪律同 MergeHeadersForm：只做参数收集与结果呈现，一切 COM 通过事件
    // 抛回 Connect 在 Ribbon 宿主线程执行（不继承 FileToolForm，不开后台
    // STA——跨线程调 WPS COM 是项目红线）。
    //
    // 一个窗体类服务两个功能区按钮（elimination.receivable /
    // elimination.revenuecost），用 EliminationKind 区分标题、提示与业务。
    //
    // 交互：来源区（工作簿/工作表/表头行）→「读取表头并预检」灌列映射网格
    // （角色按表头文本自动猜，用户可改）→ 再点预检读数据算分录 → 分录预览
    // 网格 →「执行写入」前二次确认 + 备份 + 快照；写入后「取消操作」还原
    // （新建表整表删除 / 已存在表快照回写），「确认操作」结束本批。
    // =====================================================================

    /// <summary>从界面收集的抵消请求（纯数据，不含 COM）。</summary>
    internal sealed class EliminationRequest
    {
        internal EliminationKind Kind = EliminationKind.Receivable;
        internal string Path = "";
        internal string SheetName = "";
        internal int HeaderRow = 1;
        internal List<EliminationColumn> Columns = new List<EliminationColumn>();
        internal string TargetSheetItem = EliminationGateway.NewSheetItem;
        /// <summary>往来抵消「差异表」目标（下拉项：新建或指定工作表名）。</summary>
        internal string DiffSheetItem = EliminationGateway.NewSheetItem;
        internal string StartCell = "A1";
        /// <summary>界面层就能发现的输入错误（表头行/起始单元格非法）；null 表示通过。</summary>
        internal string Error = null;

        // ---- 矩阵式底稿（收入成本抵消）专用 ----
        /// <summary>条件列的期望值，如「是」；只有该值的行参与抵消。</summary>
        internal string ConditionValue = "是";
        /// <summary>降级：写静态数值而非活公式。</summary>
        internal bool StaticValues;
        /// <summary>引用整列（$O:$O）；关闭则限定源数据区。</summary>
        internal bool FullColumnRefs = true;
        /// <summary>首行写 SUBTOTAL 合计。</summary>
        internal bool TotalRow = true;
        /// <summary>自动探测表头行（源表首行是汇总行时下移一行）。</summary>
        internal bool AutoHeader = true;
        /// <summary>每个交易对之间空几行（0=不空行）。</summary>
        internal int GapRows;
        internal EliminationRequest Copy()
        {
            var copy=(EliminationRequest)MemberwiseClone();
            copy.Columns=new List<EliminationColumn>();
            foreach(var column in Columns)copy.Columns.Add(new EliminationColumn {Number=column.Number,Letter=column.Letter,Header=column.Header,Role=column.Role});
            return copy;
        }
    }

    internal sealed class EliminationForm : Form, IDialogDefaultLayout
    {
        // ---- 来源区 ----
        private readonly TextBox sourcePath = new TextBox { Width = 420, Name = "sourceWorkbook" };
        private readonly ComboBox sourceSheet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Width = 190, Name = "sourceSheetBox" };
        private readonly TextBox headerRow = new TextBox { Width = 60, Text = "2", Name = "headerRowBox" };
        // 真实底稿常在第 1 行放 SUBTOTAL 汇总行，表头实际在第 2 行，故默认 2。
        private readonly TextBox conditionValue = new TextBox { Width = 60, Text = "是", Name = "conditionValueBox" };
        // ---- 目标区 ----
        private readonly ComboBox targetSheet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Width = 190, Name = "targetSheetBox" };
        /// <summary>往来抵消「差异表」目标下拉（收入成本抵消不显示）。</summary>
        private readonly ComboBox diffSheet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Width = 190, Name = "diffSheetBox" };
        private readonly TextBox startCell = new TextBox { Width = 80, Text = "A1", Name = "startCellBox" };
        // ---- 矩阵式底稿选项（仅收入成本抵消显示）；AutoSize 防止高 DPI 下文本被截断 ----
        // 「写数值不写公式」复选框已删除：用户实机反馈写公式多次报错（跨表引用核对失败等），
        // 默认只写数值（MatrixOptions.StaticValues 恒 true），不再写活公式。
        private readonly CheckBox fullColumnRefs = new CheckBox { Text = "引用整列", AutoSize = true, Checked = true, Name = "fullColumnRefsBox" };
        private readonly CheckBox totalRow = new CheckBox { Text = "首行合计", AutoSize = true, Checked = true, Name = "totalRowBox" };
        private readonly NumericUpDown gapRows = new NumericUpDown
        { Width = 56, Minimum = 0, Maximum = 20, Value = 0, Name = "gapRowsBox" };

        private readonly Label hint = Ui.Label("", 9, true);
        private readonly RichTextBox logView = new RichTextBox
        { ReadOnly = true, WordWrap = true, ScrollBars = RichTextBoxScrollBars.Vertical, Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Color.White };
        private readonly FlowLayoutPanel actions = Ui.ActionBar();
        private readonly TableLayoutPanel root;
        /// <summary>列映射左半区（源表前半列）。字段名保持 Map 不变——DialogFormatting 用字段名生成 id（Map.0），改名会让用户已保存的自定义字体失效。</summary>
        internal readonly ResultGrid Map = new ResultGrid();
        /// <summary>列映射右半区（源表后半列）。与 Map 并列放在可拖动分隔条两侧，避免 26 列时要滚很久。</summary>
        internal readonly ResultGrid MapRight = new ResultGrid();
        internal readonly ResultGrid Entries = new ResultGrid();
        /// <summary>列映射容器（左右分栏、可拖动）。</summary>
        private readonly SplitContainer mapSplit;
        internal readonly RoundedButton Preview, Execute, Stop, Confirm, CopyLog, CloseButton;

        private readonly EliminationKind kind;
        private EliminationResult preview;
        /// <summary>收入成本抵消（矩阵式底稿）的预检结果；与 preview 并存，按 kind 取用。</summary>
        private EliminationMatrix previewMatrix;
        private object[][] virtualValues;
        private EliminationMatrix displayMatrix;
        private bool busy;
        private bool canUndo;
        internal Action CancelPreview;

        internal event Action PreviewRequested;
        internal event Action ExecuteRequested;
        internal event Action UndoRequested;
        internal event Action ConfirmRequested;
        /// <summary>来源工作簿变更后，请求读一次它的工作表名（COM 只能在宿主线程做）。</summary>
        internal event Action<string> SheetsRequested;

        internal EliminationForm(EliminationKind kind)
        {
            this.kind = kind;
            Entries.CellValueNeeded+=delegate(object sender,DataGridViewCellValueEventArgs cell)
            {
                if(virtualValues==null||displayMatrix==null||cell.RowIndex<0||cell.RowIndex>=virtualValues.Length)return;
                var values=virtualValues[cell.RowIndex];
                if(values==null)virtualValues[cell.RowIndex]=values=PreviewValues(displayMatrix.Rows[cell.RowIndex],displayMatrix);
                if(cell.ColumnIndex>=0&&cell.ColumnIndex<values.Length)cell.Value=values[cell.ColumnIndex];
            };
            Entries.RowHeightInfoNeeded+=delegate(object sender,DataGridViewRowHeightInfoNeededEventArgs row)
            {if(Entries.VirtualMode)row.Height=Entries.RowTemplate.Height;};
            Text = kind == EliminationKind.Receivable ? "往来抵消" : "收入成本抵消";
            DialogSizing.Prepare(this,
                kind == EliminationKind.Receivable ? "elimination.receivable" : "elimination.revenuecost",
                "table", Ui.Blue, FormWindowState.Maximized);
            KeyPreview = true;

            root = new BufferedTableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 12, Padding = new Padding(16) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 0  来源区
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 2));   // 1  间隔：来源 ↔ 提示（已无拖柄）
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 2  提示
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 3  目标区
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 2));   // 4  间隔：目标 ↔ 列映射（已无拖柄）
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 40));   // 5  列映射
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));   // 6  手柄：列映射 ↔ 分录预览
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 40));   // 7  分录预览
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));   // 8  手柄：分录 ↔ 报告
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 20));   // 9  报告
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));   // 10 手柄：报告下边界（按钮栏固定）
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 11 按钮

            // ---- 区域 1：来源 ----
            // 两种抵消业务现在都走矩阵式输出：收入成本抵消 →「收入抵消结果」；
            // 往来抵消 →「往来抵消差异」+「往来抵消结果」。故目标区选项与结果网格恒显示。
            bool matrix = true;

            // Keep each label/input/grip together, wrapping groups when DPI or restored
            // widths exceed the available client area instead of hiding trailing fields.
            FlowLayoutPanel sourceRow = new BufferedFlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, Name = "sourceRow" };
            RoundedButton pickSource = Ui.Button("选择来源…", 108);
            pickSource.Name = "pickSourceButton";
            pickSource.Click += delegate { PickFile(); };
            sourceRow.Controls.Add(Ui.Grip(sourcePath, "来源工作簿"));
            sourceRow.Controls.Add(Ui.Grip(pickSource, (Control)null));
            sourceRow.Controls.Add(Ui.Grip(sourceSheet, "工作表"));
            sourceRow.Controls.Add(Ui.Grip(headerRow, "表头行"));
            if (matrix) sourceRow.Controls.Add(Ui.Grip(conditionValue, "条件值"));

            // ---- 区域 2：目标 ----
            // 往来抵消产出两张表：差异表 + 结果表，各一个下拉，消除「一张下拉管两张表」的歧义；
            // 收入成本只有一张结果表，下拉标签仍是「写入工作表」。
            bool receivable = kind == EliminationKind.Receivable;
            FlowLayoutPanel targetRow = new BufferedFlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, Name = "targetRow" };
            if (receivable) targetRow.Controls.Add(Ui.Grip(diffSheet, "差异表"));
            targetRow.Controls.Add(Ui.Grip(targetSheet, receivable ? "结果表" : "写入工作表"));
            targetRow.Controls.Add(Ui.Grip(startCell, "起始单元格"));
            if (matrix)
            {
                targetRow.Controls.Add(Ui.Grip(fullColumnRefs, (Control)null));
                targetRow.Controls.Add(Ui.Grip(totalRow, (Control)null));
                targetRow.Controls.Add(Ui.Grip(gapRows, "写入空行"));
            }

            // 选项一改，上次预检结果即作废
            EventHandler invalidate = delegate { InvalidatePreview(); };
            fullColumnRefs.CheckedChanged += invalidate;
            totalRow.CheckedChanged += invalidate;
            gapRows.ValueChanged += invalidate;
            conditionValue.TextChanged += invalidate;

            // ---- 提示 ----
            hint.Margin = new Padding(0, 2, 0, 2);
            hint.MaximumSize = new Size(4096, TextRenderer.MeasureText("测", hint.Font, Size.Empty,
                TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding).Height + 2);
            // hint 被锁成单行高，只能用短句；完整口径说明走 SetLog / AppendLog。
            hint.Text = kind == EliminationKind.Receivable
                ? "按（核算主体,对方单位,年度）双向合计后以小抵大：产出「往来抵消差异」（每主体对 3 个科目对）与「往来抵消结果」两张表；年度取截止日前 4 位。"
                : "按（核算主体,客商,年度）三维取值出矩阵底稿：收入借、成本贷、成本还原贷、存货贷；U 优先取损益列，无则按(收入−成本)×比例现算。";

            // ---- 网格 ----
            // 列映射左右分栏：源表列多时（如 26 列）避免向下滚动很久；分隔条可拖动。
            // Map 先构造、先入控件树，保证 DialogFormatting.Walk 取到的 grids[0] 仍是左半区（GridHeadersProbe 依赖）。
            BuildMapGrid(Map);
            BuildMapGrid(MapRight);
            // 矩阵式底稿列数随年度变化，等预检出结果再重建（见 BuildResultGrid）；
            // 往来抵消仍是固定 6 列，构造时一次建好。
            if (kind == EliminationKind.Receivable) BuildEntriesGrid();
            else BuildEntriesBase();
            // 注意：构造时**不要**设 Panel1MinSize/Panel2MinSize——SplitContainer 宽度此时还是默认值，
            // 设较大的 MinSize 会在内部触发 SplitterDistance 越界异常（离线预览构造时就崩）。
            // 两个最小宽度改由 CenterMapSplitter() 在布局完成后设置。
            mapSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                SplitterWidth = 8,
                Name = "mapSplit"
            };
            mapSplit.Panel1.Controls.Add(Map);
            mapSplit.Panel2.Controls.Add(MapRight);
            UiCard mapCard = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(8), Margin = new Padding(0, 0, 0, 8) };
            mapCard.Controls.Add(mapSplit);
            UiCard entriesCard = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(8), Margin = new Padding(0, 0, 0, 8) };
            entriesCard.Controls.Add(Entries);

            // ---- 报告 ----
            UiCard logCard = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(8), Margin = new Padding(0, 0, 0, 8) };
            logCard.Controls.Add(logView);

            // ---- 按钮栏 ----
            Preview = Ui.Button("读取表头并预检", 140); Ui.StylePrimary(Preview);
            Execute = Ui.Button("执行写入", 108);
            Stop = Ui.Button("取消操作", 104);
            Confirm = Ui.Button("确认操作", 104);
            CopyLog = Ui.Button("复制日志", 100);
            CloseButton = Ui.Button("关闭", 88);
            actions.Controls.AddRange(new Control[] { Preview, Execute, Stop, Confirm, CopyLog, CloseButton });
            ToolTip tip = new ToolTip();
            tip.SetToolTip(fullColumnRefs, "勾选（默认）：公式引用整列/算到工作表末尾，源表新增行会自动纳入；"
                + "取消：限定到当前数据区，重算更快，但源表新增行后需重新执行才能纳入。");
            tip.SetToolTip(Preview, "读取来源表头、按角色读取数据并计算抵消分录（Ctrl+Shift+P）");
            tip.SetToolTip(Execute, "把分录写入目标工作表（Ctrl+Shift+W）");
            tip.SetToolTip(Stop, "还原：新建表整表删除；已存在表用写入前快照回写");
            tip.SetToolTip(Confirm, "确认本次结果，结束可还原状态（文件备份保留）");

            Preview.Click += delegate { if (PreviewRequested != null) PreviewRequested(); };
            Execute.Click += delegate { if (ExecuteRequested != null) ExecuteRequested(); };
            Stop.Click += delegate { if (busy) { if(CancelPreview!=null){CancelPreview();return;} UiDialog.Show(this, "正在处理，请等待当前步骤结束。", Text, MessageBoxButtons.OK, MessageBoxIcon.Information); return; } if (UndoRequested != null) UndoRequested(); };
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

            // 各方框下方一条 8px 拖拽条；报告框（最下方）只改自身高度，按钮栏固定。
            // key 带 .v2 版本后缀：列映射改左右分栏 + 弹窗改为打开即最大化后，
            // 旧存档（row.map=371 / entries=336 / report=494，合计远超窗口高度、把按钮栏挤出可视区）
            // 不再适用，升版本使其自然失效、回到 Percent 按比例分配。
            //
            // 来源区（行 0）与目标区（行 3）**不设拖拽条**：这两行是 AutoSize 内容行，
            // 拖动只会把它们钉成 Absolute（实测被顶到 minAbove=92，内容仅需 ~55px），
            // 卡片随 Dock=Fill 被拉高而内容贴顶 → 卡片下方凭空留白。拖它对用户无正向价值。
            // 行 1 / 行 4 的 8px 仍保留为纯视觉间隔。
            RowGrip gripMap = new RowGrip(root, 5, 7, "row.map.v2", 120, 96);
            RowGrip gripEntries = new RowGrip(root, 7, 9, "row.entries.v2", 120, 120);
            RowGrip gripReport = new RowGrip(root, 9, 11, "row.report.v2", 120, 0, true);

            // 关键：AutoSize 行里放 Dock=Fill 的子控件，TLP 会读子控件「当前实际高度」而不是期望高度（循环依赖），
// 卡片永远缩不回去——这就是多轮压缩 Padding/Margin 都无效的根因。改 Dock=Top + AutoSize 让高度由内容决定。
            UiCard sourceCard = new FieldRowCard(sourceRow) { AutoSize = true, Dock = DockStyle.Top, Padding = new Padding(8, 4, 8, 4), Margin = new Padding(0, 0, 0, 4) };
            sourceCard.Controls.Add(sourceRow);
            UiCard targetCard = new FieldRowCard(targetRow) { AutoSize = true, Dock = DockStyle.Top, Padding = new Padding(8, 4, 8, 4), Margin = new Padding(0, 0, 0, 4) };
            targetCard.Controls.Add(targetRow);

            root.Controls.Add(sourceCard, 0, 0);
            root.Controls.Add(hint, 0, 2);
            root.Controls.Add(targetCard, 0, 3);
            root.Controls.Add(mapCard, 0, 5);
            root.Controls.Add(gripMap, 0, 6);
            root.Controls.Add(entriesCard, 0, 7);
            root.Controls.Add(gripEntries, 0, 8);
            root.Controls.Add(logCard, 0, 9);
            root.Controls.Add(gripReport, 0, 10);
            root.Controls.Add(actions, 0, 11);
            Controls.Add(root);

            RefreshActions();
        }

        internal void ClearPreview() {preview=null;previewMatrix=null;RefreshActions();}

        private sealed class FieldRowCard : UiCard
        {
            private readonly FlowLayoutPanel fields;
            internal FieldRowCard(FlowLayoutPanel row) {fields=row;}
            public override Size GetPreferredSize(Size proposedSize)
            {
                // Panel's unbounded preferred size assumes one line. Measure at
                // the allocated width so its parent reserves every wrapped row.
                int width=proposedSize.Width>0&&proposedSize.Width<int.MaxValue?proposedSize.Width:Width;
                var size=fields.GetPreferredSize(new Size(Math.Max(1,width-Padding.Horizontal),0));
                return new Size(width,size.Height+Padding.Vertical);
            }
        }

        // ------------------------ 网格 ------------------------

        /// <summary>构造一个列映射网格（左右两格共用同一套列定义：列 / 表头列名 / 角色）。</summary>
        private void BuildMapGrid(ResultGrid grid)
        {
            grid.Dock = DockStyle.Fill;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.RowHeadersVisible = false;
            grid.MultiSelect = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.AutoGenerateColumns = false;
            grid.ScrollBars = ScrollBars.Both;
            grid.RowTemplate.Height = 28;
            grid.EnableHeadersVisualStyles = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            grid.ShowCellToolTips = true;
            MakeColumn(grid, "letter", "列", 56, 90, true);
            MakeColumn(grid, "header", "表头列名", 160, 260, true);
            DataGridViewComboBoxColumn role = new DataGridViewComboBoxColumn
            { Name = "role", HeaderText = "角色（列含义）", Width = 150, SortMode = DataGridViewColumnSortMode.NotSortable };
            // 只装本业务真正用到的角色：收入成本抵消走矩阵式，角色集合与往来抵消不同。
            foreach (EliminationRole item in MatrixRoleList.For(kind))
                role.Items.Add(EliminationEngine.RoleText(item));
            grid.Columns.Add(role);
            grid.Columns[grid.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            // 下拉列 Items 与渲染示例值不匹配时会弹模态 DataError，离线探针会挂起——必须吞掉。
            grid.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            // 角色是猜的：用户一改就意味着上次预检结果作废，回到「需重新预检」。
            grid.CellValueChanged += delegate { InvalidatePreview(); };
            HookComboOneClick(grid);
            Ui.StyleGrid(grid);
        }

        /// <summary>
        /// 角色下拉「单击即展开」：默认 EditMode 下首次点击只是选中单元格、第二次点击才展开，
        /// 用户要连点两次。这里在 CellClick 时直接进入编辑并展开下拉。
        /// 只作用于下拉列，不改全局 EditMode（否则文本框列也会变成单击即编辑）。
        /// </summary>
        private static void HookComboOneClick(ResultGrid grid)
        {
            grid.CellClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                if (grid.ReadOnly) return;
                if (!(grid.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn)) return;
                grid.CurrentCell = grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                grid.BeginEdit(true);
                DataGridViewComboBoxEditingControl combo =
                    grid.EditingControl as DataGridViewComboBoxEditingControl;
                if (combo != null) combo.DroppedDown = true;
            };
        }

        /// <summary>
        /// 结果网格的公共属性与表头样式。矩阵式底稿的列要等预检后才知道，
        /// 但表头样式必须在构造时就套好（离线预览探针会断言这一点）。
        /// </summary>
        private void BuildEntriesBase()
        {
            Entries.Dock = DockStyle.Fill;
            Entries.AllowUserToAddRows = false;
            Entries.AllowUserToDeleteRows = false;
            Entries.AllowUserToResizeRows = false;
            Entries.RowHeadersVisible = false;
            Entries.MultiSelect = false;
            Entries.ReadOnly = true;
            Entries.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            Entries.AutoGenerateColumns = false;
            Entries.ScrollBars = ScrollBars.Both;
            Entries.RowTemplate.Height = 28;
            Entries.EnableHeadersVisualStyles = false;
            Entries.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            Entries.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            Entries.ShowCellToolTips = true;
            Entries.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            Ui.StyleGrid(Entries);
        }

        private void BuildEntriesGrid()
        {
            BuildEntriesBase();
            MakeColumn(Entries, "index", EliminationGateway.EntryHeaders[0], 80, 100, true);
            MakeColumn(Entries, "summary", EliminationGateway.EntryHeaders[1], 220, 340, true);
            MakeColumn(Entries, "direction", EliminationGateway.EntryHeaders[2], 60, 80, true);
            MakeColumn(Entries, "account", EliminationGateway.EntryHeaders[3], 140, 200, true);
            DataGridViewTextBoxColumn amount = new DataGridViewTextBoxColumn
            { Name = "amount", HeaderText = EliminationGateway.EntryHeaders[4], Width = 120, SortMode = DataGridViewColumnSortMode.NotSortable };
            amount.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            amount.DefaultCellStyle.Format = "0.00";
            Entries.Columns.Add(amount);
            MakeColumn(Entries, "note", EliminationGateway.EntryHeaders[5], 260, 0, true);
            Entries.Columns[Entries.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        // 列宽上限存在 Tag，由 Ui.AutoFitColumns 读取并封顶；最后一列 Fill 由调用方设置。
        private static DataGridViewTextBoxColumn MakeColumn(DataGridView grid, string name, string header, int width, int cap, bool notSortable)
        {
            DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
            column.Name = name;
            column.HeaderText = header;
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            column.MinimumWidth = 50;
            if (notSortable) column.SortMode = DataGridViewColumnSortMode.NotSortable;
            Font headerFont = grid.ColumnHeadersDefaultCellStyle.Font ?? grid.Font;
            int needed = TextRenderer.MeasureText(header, headerFont).Width + 32;
            column.Width = Math.Max(width, needed);
            if (cap > 0) column.Tag = Math.Max(cap, needed);
            grid.Columns.Add(column);
            return column;
        }

        // ------------------------ 数据 ------------------------

        /// <summary>从界面收集请求（不含 COM）。角色列当前值实时收进 Columns。</summary>
        internal EliminationRequest BuildRequest()
        {
            EliminationRequest request = new EliminationRequest();
            request.Kind = kind;
            request.Path = sourcePath.Text == null ? "" : sourcePath.Text.Trim();
            string sheet = Convert.ToString(sourceSheet.Text);
            request.SheetName = sheet == null ? "" : sheet.Trim();
            int header;
            if (!int.TryParse(headerRow.Text == null ? "" : headerRow.Text.Trim(),
                    NumberStyles.None, CultureInfo.InvariantCulture, out header) || header < 1)
                request.Error = "表头行须为不小于 1 的整数。";
            else request.HeaderRow = header;
            request.TargetSheetItem = Convert.ToString(targetSheet.Text);
            if (string.IsNullOrEmpty(request.TargetSheetItem)) request.TargetSheetItem = EliminationGateway.NewSheetItem;
            if (kind == EliminationKind.Receivable)
            {
                request.DiffSheetItem = Convert.ToString(diffSheet.Text);
                if (string.IsNullOrEmpty(request.DiffSheetItem)) request.DiffSheetItem = EliminationGateway.NewSheetItem;
            }
            request.StartCell = startCell.Text == null ? "" : startCell.Text.Trim();
            if (request.StartCell.Length == 0) request.StartCell = "A1";

            request.ConditionValue = conditionValue.Text == null ? "" : conditionValue.Text.Trim();
            if (request.ConditionValue.Length == 0) request.ConditionValue = "是";
            // 「写数值不写公式」复选框已删除：恒写数值（写公式多次在宿主核对环节报错）
            request.StaticValues = true;
            request.FullColumnRefs = fullColumnRefs.Checked;
            request.TotalRow = totalRow.Checked;
            request.GapRows = (int)gapRows.Value;
            request.AutoHeader = true;

            // 先左半区后右半区，保持源表 A→Z 的列顺序（签名与校验都依赖这个顺序）
            foreach (ResultGrid grid in new[] { Map, MapRight })
            {
                foreach (DataGridViewRow row in grid.Rows)
                {
                    EliminationColumn column = row.Tag as EliminationColumn;
                    if (column == null) continue;
                    column.Role = RoleFromText(Convert.ToString(row.Cells["role"].Value));
                    request.Columns.Add(column);
                }
            }
            return request;
        }

        /// <summary>仅收集当前网格的列角色（不触发 BuildRequest 的其它副作用），供关窗保存。</summary>
        internal List<EliminationColumn> CollectColumns()
        {
            List<EliminationColumn> columns = new List<EliminationColumn>();
            foreach (ResultGrid grid in new[] { Map, MapRight })
            {
                foreach (DataGridViewRow row in grid.Rows)
                {
                    EliminationColumn column = row.Tag as EliminationColumn;
                    if (column == null) continue;
                    column.Role = RoleFromText(Convert.ToString(row.Cells["role"].Value));
                    columns.Add(column);
                }
            }
            return columns;
        }

        /// <summary>关窗时无条件保存当前列映射（临时记忆，按文件+业务键）。</summary>
        internal void PersistMapping()
        {
            try
            {
                string path = sourcePath.Text == null ? "" : sourcePath.Text.Trim();
                if (string.IsNullOrEmpty(path)) return;
                List<EliminationColumn> columns = CollectColumns();
                if (columns.Count == 0) return;
                EliminationMapStore.Save(EliminationMapStore.Key(kind, path), columns);
            }
            catch { }
        }

        private EliminationRole RoleFromText(string text)
        {
            if (string.IsNullOrEmpty(text)) return EliminationRole.Ignore;
            foreach (EliminationRole role in MatrixRoleList.For(kind))
                if (EliminationEngine.RoleText(role) == text) return role;
            return EliminationRole.Ignore;
        }

        /// <summary>各业务可用的列角色清单（列映射下拉与回读共用，避免两处不一致）。</summary>
        internal static class MatrixRoleList
        {
            internal static EliminationRole[] For(EliminationKind kind)
            {
                if (kind == EliminationKind.Receivable)
                    return new[]
                    {
                        EliminationRole.Ignore, EliminationRole.Entity, EliminationRole.Counterparty,
                        EliminationRole.Account, EliminationRole.Amount, EliminationRole.Deadline,
                        EliminationRole.Condition, EliminationRole.BadDebt
                    };
                return new[]
                {
                    EliminationRole.Ignore, EliminationRole.Entity, EliminationRole.Counterparty,
                    EliminationRole.Year, EliminationRole.Condition,
                    EliminationRole.Revenue, EliminationRole.Cost,
                    EliminationRole.UnsoldRatio, EliminationRole.UnrealizedProfit
                };
            }
        }

        /// <summary>
        /// 列映射临时记忆：按「文件路径 + 表名 + 业务种类」为键，内存保存用户上次指定的角色，
        /// 下次对同一活动工作簿打开弹窗自动带出；完全关闭工作簿后调用 ClearFile 清除，
        /// 避免 XML 落盘持续膨胀。
        /// </summary>
        internal static class EliminationMapStore
        {
            // 键 = 文件路径|业务种类|表名 → (列字母 → 角色文本)
            private static readonly Dictionary<string, Dictionary<string, string>> Memory =
                new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);

            /// <summary>记忆键：文件路径|业务种类。</summary>
            internal static string Key(EliminationKind kind, string path)
            {
                return (path ?? "") + "|" + kind;
            }

            /// <summary>命中返回「列字母 → 角色文本」；未命中返回 null。</summary>
            internal static Dictionary<string, string> Load(string key)
            {
                Dictionary<string, string> map;
                return Memory.TryGetValue(key, out map) ? map : null;
            }

            /// <summary>保存该来源的列映射（含忽略角色，记住用户把某列改为忽略）。</summary>
            internal static void Save(string key, List<EliminationColumn> columns)
            {
                if (columns == null) return;
                Dictionary<string, string> mine = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (EliminationColumn column in columns)
                    mine[column.Letter] = EliminationEngine.RoleText(column.Role);
                if (mine.Count == 0) return;
                Memory[key] = mine;
            }

            /// <summary>关闭工作簿时清除该文件路径下的所有记忆条目。</summary>
            internal static void ClearFile(string path)
            {
                if (string.IsNullOrEmpty(path)) return;
                string prefix = path + "|";
                List<string> remove = new List<string>();
                foreach (string key in Memory.Keys)
                    if (key.StartsWith(prefix, StringComparison.Ordinal)) remove.Add(key);
                foreach (string key in remove) Memory.Remove(key);
            }

            /// <summary>懒清理：清除不在给定打开文件集合内的记忆条目（已关闭的工作簿）。</summary>
            internal static void CleanupExcept(HashSet<string> openPaths)
            {
                if (openPaths == null) return;
                List<string> remove = new List<string>();
                foreach (string key in Memory.Keys)
                {
                    int sep = key.IndexOf('|');
                    if (sep < 0) continue;
                    if (!openPaths.Contains(key.Substring(0, sep))) remove.Add(key);
                }
                foreach (string key in remove) Memory.Remove(key);
            }
        }

        private void PickFile()
        {
            if (busy || canUndo) return;
            if (DialogPreviewFactory.IsConstructing) return;
            string[] selected = HostDialogService.OpenFiles(this);
            if (selected == null || selected.Length == 0) return;
            sourcePath.Text = selected[0];
            sourceSheet.Items.Clear();
            sourceSheet.Text = "";
            targetSheet.Items.Clear();
            targetSheet.Text = EliminationGateway.NewSheetItem;
            preview = null;
            Map.Rows.Clear();
            MapRight.Rows.Clear();
            Entries.Rows.Clear();
            if (SheetsRequested != null) SheetsRequested(selected[0]);
        }

        /// <summary>Connect 在宿主线程取到当前活动工作簿后回填来源区。</summary>
        internal void SetSource(string path, string sheet)
        {
            if (IsDisposed) return;
            sourcePath.Text = path == null ? "" : path;
            if (!string.IsNullOrEmpty(sheet)) sourceSheet.Text = sheet;
            if (SheetsRequested != null) SheetsRequested(path);
        }

        /// <summary>回写来源/目标两个工作表下拉（目标下拉末尾多一项「（新建工作表）」）。</summary>
        internal void SetSheets(List<string> names, string error)
        {
            if (IsDisposed) return;
            if (!string.IsNullOrEmpty(error)) { AppendLog("读取工作表失败：" + error); return; }
            string keepSource = Convert.ToString(sourceSheet.Text);
            string keepTarget = Convert.ToString(targetSheet.Text);
            string keepDiff = Convert.ToString(diffSheet.Text);
            sourceSheet.Items.Clear();
            if (names != null) foreach (string name in names) sourceSheet.Items.Add(name);
            if (names != null && names.Count > 0)
                sourceSheet.Text = names.Contains(keepSource) ? keepSource : names[0];

            targetSheet.Items.Clear();
            if (names != null) foreach (string name in names) targetSheet.Items.Add(name);
            targetSheet.Items.Add(EliminationGateway.NewSheetItem);
            bool keep = keepTarget == EliminationGateway.NewSheetItem
                || (names != null && names.Contains(keepTarget));
            targetSheet.Text = keep && !string.IsNullOrEmpty(keepTarget)
                ? keepTarget : EliminationGateway.NewSheetItem;

            // 往来抵消的「差异表」下拉与结果表同源；收入成本不显示也不收集
            if (kind == EliminationKind.Receivable)
            {
                diffSheet.Items.Clear();
                if (names != null) foreach (string name in names) diffSheet.Items.Add(name);
                diffSheet.Items.Add(EliminationGateway.NewSheetItem);
                bool keepD = keepDiff == EliminationGateway.NewSheetItem
                    || (names != null && names.Contains(keepDiff));
                diffSheet.Text = keepD && !string.IsNullOrEmpty(keepDiff)
                    ? keepDiff : EliminationGateway.NewSheetItem;
            }
            RefreshActions();
        }

        /// <summary>
        /// 表头读入后建列映射网格（角色按表头文本猜，用户可改）。
        /// 若本文件+表名与上次保存的一致，自动带出用户上次指定的角色（临时记忆，关闭工作簿后清除）。
        /// </summary>
        internal void SetHeaders(List<EliminationColumn> columns, int dataRows, string error)
        {
            if (IsDisposed) return;
            Map.Rows.Clear();
            MapRight.Rows.Clear();
            if (!string.IsNullOrEmpty(error)) { SetLog("读取表头失败：" + error); return; }
            if (columns == null || columns.Count == 0) { SetLog("表头行没有可识别的列。"); return; }

            // 左右分栏：前半放左格、后半放右格（26 列即左 A-N、右 O-Z）
            int half = (columns.Count + 1) / 2;
            string key = EliminationMapStore.Key(kind, sourcePath.Text);
            Dictionary<string, string> saved = EliminationMapStore.Load(key);
            for (int i = 0; i < columns.Count; i++)
            {
                EliminationColumn column = columns[i];
                if (saved != null)
                {
                    string roleText;
                    if (saved.TryGetValue(column.Letter, out roleText))
                        column.Role = RoleFromText(roleText);
                }
                ResultGrid grid = i < half ? Map : MapRight;
                int index = grid.Rows.Add(column.Letter, column.Header, EliminationEngine.RoleText(column.Role));
                grid.Rows[index].Tag = column;
            }
            Ui.AutoFitColumns(Map);
            Ui.AutoFitColumns(MapRight);
            if (saved != null)
            {
                SetLog("已自动载入上次保存的列映射配置（同一文件+表名）。\r\n"
                    + "请核对角色是否正确；如需调整，改完角色后重新点「读取表头并预检」，"
                    + "预检通过后才能执行写入。\r\n");
            }
            else
            {
                SetLog("已读入 " + columns.Count + " 列表头；数据区 " + dataRows + " 行。"
                    + "请为各列选择角色（首次使用需逐列指定；保存后再次选择同一文件+表名会自动带出，"
                    + "关闭工作簿后记忆清除），然后点「读取表头并预检」。");
            }
        }

        /// <summary>预检完成：灌分录预览网格与报告。</summary>
        internal void SetPreview(EliminationResult result)
        {
            if (IsDisposed) return;
            preview = result;
            Entries.Rows.Clear();
            if (result == null) { RefreshActions(); return; }

            foreach (EliminationEntry entry in result.Entries)
                foreach (EliminationLine line in entry.Lines)
                    Entries.Rows.Add(entry.Index, entry.Summary, line.IsDebit ? "借" : "贷",
                        line.Account, line.Amount, line.Note);

            StringBuilder text = new StringBuilder();
            foreach (string error in result.Errors) text.AppendLine("阻断：" + error);
            for (int i = 0; i < result.Warnings.Count && i < 30; i++)
                text.AppendLine("提示：" + result.Warnings[i]);
            if (result.Warnings.Count > 30)
                text.AppendLine("…等共 " + result.Warnings.Count + " 条提示");
            if (result.Errors.Count == 0)
            {
                text.AppendLine("共 " + result.Entries.Count + " 条分录；总借 "
                    + result.TotalDebit.ToString("0.00", CultureInfo.InvariantCulture)
                    + "、总贷 " + result.TotalCredit.ToString("0.00", CultureInfo.InvariantCulture)
                    + (result.Balanced ? "（借贷平衡）" : "（⚠ 借贷不平，请检查）"));
                if (result.SkippedRows > 0) text.AppendLine("跳过 " + result.SkippedRows + " 行（原因见上方提示）。");
                if (result.Entries.Count > 0) text.AppendLine("请点「执行写入」。");
                else text.AppendLine("没有可写入的分录。");
            }
            SetLog(text.ToString());
            RefreshActions();
        }

        /// <summary>矩阵式底稿的结果网格：年度列数随数据变化，每次预检重建。</summary>
        private void BuildResultGrid(EliminationMatrix matrix)
        {
            Entries.Columns.Clear();
            MakeColumn(Entries, "entity", "核算主体", 120, 160, true);
            MakeColumn(Entries, "customer", "明细【客商】", 140, 200, true);
            MakeColumn(Entries, "note", "备注", 220, 320, true);
            MakeColumn(Entries, "adjust", "调整事项", 110, 140, true);
            MakeColumn(Entries, "account", "会计科目", 100, 140, true);
            if (matrix != null)
            {
                foreach (MatrixColumn column in matrix.Columns)
                {
                    if (column.IsSpacer)
                    {
                        DataGridViewTextBoxColumn spacer = MakeColumn(Entries, "sp" + column.Index, "", 18, 18, true);
                        spacer.DefaultCellStyle.BackColor = Color.FromArgb(0xF2, 0xF2, 0xF2);
                        continue;
                    }
                    DataGridViewTextBoxColumn money = new DataGridViewTextBoxColumn
                    {
                        Name = "y" + column.Index,
                        HeaderText = column.Year + (column.IsDebit ? "·借" : "·贷"),
                        Width = 130,
                        SortMode = DataGridViewColumnSortMode.NotSortable
                    };
                    money.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    money.DefaultCellStyle.Format = "0.00";
                    Entries.Columns.Add(money);
                }
            }
            Entries.Columns[Entries.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Ui.StyleGrid(Entries);
        }

        /// <summary>预检完成（收入成本抵消）：重建结果网格并灌行与报告。</summary>
        internal void SetPreviewMatrix(EliminationMatrix matrix)
        {
            if (IsDisposed) return;
            Entries.Rows.Clear();Entries.VirtualMode=false;virtualValues=null;
            previewMatrix = matrix;
            displayMatrix = matrix;
            BuildResultGrid(matrix);
            Entries.Rows.Clear();
            if (matrix == null) { RefreshActions(); return; }

            if(matrix.Rows.Count>=5000)
            {
                // Materialize only values, not one native-cell object per result cell.
                // Full-data AutoFit and native copy still obtain every value through
                // CellValueNeeded; the source matrix remains the stable row identity.
                var handle=Entries.Handle;
                virtualValues=new object[matrix.Rows.Count][];Entries.VirtualMode=true;Entries.RowCount=matrix.Rows.Count;
            }
            else
            {
                var batch=new List<DataGridViewRow>();
                foreach(MatrixRow row in matrix.Rows)
                {
                    var display=new DataGridViewRow();display.CreateCells(Entries,PreviewValues(row,matrix));display.Tag=row;batch.Add(display);
                }
                Entries.SuspendLayout();try {if(batch.Count>0)Entries.Rows.AddRange(batch.ToArray());}finally {Entries.ResumeLayout(false);}
            }
            Ui.AutoFitColumns(Entries);

            StringBuilder text = new StringBuilder();
            foreach (string error in matrix.Errors) text.AppendLine("阻断：" + error);
            int shown = Math.Min(matrix.Warnings.Count, 30);
            for (int i = 0; i < shown; i++) text.AppendLine("提示：" + matrix.Warnings[i]);
            if (matrix.Warnings.Count > 30) text.AppendLine("…等共 " + matrix.Warnings.Count + " 条提示");
            if (matrix.Errors.Count == 0)
            {
                text.AppendLine("交易对 " + matrix.GroupCount + " 组 × 年度 " + matrix.Years.Count
                    + " 个 → " + matrix.Rows.Count + " 行、" + matrix.ColumnCount + " 列"
                    + (matrix.Balanced ? "（借贷平衡）" : "（⚠ 借贷不平，请检查）"));
                text.AppendLine("年度：" + string.Join("、", matrix.Years.ToArray()));
                if (matrix.SkippedRows > 0) text.AppendLine("跳过 " + matrix.SkippedRows + " 行（原因见上方提示）。");
                if (matrix.FilteredRows > 0) text.AppendLine("条件列筛掉 " + matrix.FilteredRows + " 行。");
                if (matrix.Rows.Count > 0) text.AppendLine("请点「执行写入」。");
                else text.AppendLine("没有可写入的内容。");
            }
            SetLog(text.ToString());
            RefreshActions();
        }

        private static object[] PreviewValues(MatrixRow row,EliminationMatrix matrix)
        {
            object[] values=new object[5+matrix.Columns.Count];
            values[0]=row.Entity;values[1]=row.Customer;values[2]=row.Note;values[3]=row.Adjustment;values[4]=row.Account;
            foreach(string year in matrix.Years)
            {
                var debit=matrix.DebitColumn(year);var credit=matrix.CreditColumn(year);MatrixAmount amount;
                if(debit!=null&&row.Debit.TryGetValue(year,out amount))values[5+debit.Index]=amount.Value;
                if(credit!=null&&row.Credit.TryGetValue(year,out amount))values[5+credit.Index]=amount.Value;
            }
            return values;
        }

        /// <summary>自动探测表头行后回写输入框。</summary>
        internal void SetHeaderRow(int row)
        { headerRow.Text = row.ToString(CultureInfo.InvariantCulture); }

        /// <summary>输入变更后作废上次预检结果。</summary>
        private void InvalidatePreview()
        {
            preview = null;
            previewMatrix = null;
            RefreshActions();
        }

        /// <summary>
        /// 是否已具备可写入的预检结果。两种业务现在都产出矩阵式结果
        /// （往来抵消的「差异表」由 diff 承载，「结果表」由 previewMatrix 承载），
        /// 故统一以「结果表有内容且无阻断错误」为准。
        /// </summary>
        private bool HasWritePreview()
        {
            return previewMatrix != null && previewMatrix.Rows.Count > 0 && previewMatrix.Errors.Count == 0;
        }

        internal void SetStatus(string text)
        {
            if (IsDisposed) return;
            if(!InvokeRequired){hint.Text=text;return;}
            if(IsHandleCreated)BeginInvoke((MethodInvoker)delegate {if(!IsDisposed)hint.Text=text;});
        }

        internal void AppendLog(string text)
        {
            if (IsDisposed) return;
            // 窗体句柄未建时（SetSource → SheetsRequested 链在 Show 之前同步执行），
            // BeginInvoke 无法排队，日志会被静默丢掉——先缓存，OnHandleCreated 再刷出。
            if (!IsHandleCreated) { pendingLog.AppendLine(DateTime.Now.ToString("HH:mm:ss") + "  " + text); return; }
            BeginInvoke((MethodInvoker)delegate
            { logView.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + text + "\r\n"); });
        }

        /// <summary>句柄创建前累积的日志（修复「打开窗体之前的错误被静默吞掉」）。</summary>
        private readonly StringBuilder pendingLog = new StringBuilder();

        internal void SetLog(string text)
        { logView.Clear(); logView.AppendText(text); }

        internal void SetBusy(bool value)
        { busy = value; RefreshActions(); }

        internal void MarkUndone()
        { canUndo = false; preview = null; previewMatrix = null; Entries.Rows.Clear(); Entries.VirtualMode=false;virtualValues=null;displayMatrix=null;RefreshActions(); }

        internal bool IsBusy { get { return busy; } }
        internal bool CanUndo { get { return canUndo; } }

        /// <summary>窗体服务的业务种类（两个功能区按钮共用一个窗体槽位时区分）。</summary>
        internal EliminationKind Kind { get { return kind; } }

        /// <summary>
        /// 列映射的总行数（左右两格之和）。列映射拆成左右两格后，
        /// 判断「网格是否为空」必须用它——只看左格会在列数很少时误判。
        /// </summary>
        internal int MapRowCount { get { return Map.Rows.Count + MapRight.Rows.Count; } }

        /// <summary>
        /// 分隔条居中。必须钳制在合法区间内——容器尚未布局完成时宽度可能很小，
        /// 直接赋 Width/2 会抛「SplitterDistance 必须在 Panel1MinSize 和 Width-Panel2MinSize 之间」。
        /// </summary>
        private void CenterMapSplitter()
        {
            if (mapSplit == null) return;
            int width = mapSplit.Width;
            if (width <= 0) return;
            // 只在宽度足够时才抬高最小宽度；窄窗口（如探针把窗体缩到最小尺寸）保持默认 25
            int want = 160;
            if (width >= want * 2 + mapSplit.SplitterWidth)
            {
                try { mapSplit.Panel1MinSize = want; mapSplit.Panel2MinSize = want; } catch { }
            }
            int min = mapSplit.Panel1MinSize;
            int max = width - mapSplit.Panel2MinSize - mapSplit.SplitterWidth;
            if (max <= min) return;
            int distance = width / 2;
            if (distance < min) distance = min;
            if (distance > max) distance = max;
            try { mapSplit.SplitterDistance = distance; } catch { }
        }

        private void RefreshActions()
        {
            bool idle = !busy && !canUndo;
            // 两种抵消业务现在都走矩阵式输出：收入成本抵消 →「收入抵消结果」；
            // 往来抵消 →「往来抵消差异」+「往来抵消结果」。故目标区选项与结果网格恒显示。
            bool matrix = true;
            sourcePath.Enabled = idle;
            sourceSheet.Enabled = idle;
            headerRow.Enabled = idle;
            targetSheet.Enabled = idle;
            diffSheet.Enabled = idle && kind == EliminationKind.Receivable;
            startCell.Enabled = idle;
            conditionValue.Enabled = idle && matrix;
            fullColumnRefs.Enabled = totalRow.Enabled = gapRows.Enabled = idle && matrix;
            MapRight.ReadOnly = Map.ReadOnly = !idle;
            Preview.Enabled = idle;
            // 两种业务的预检结果模型不同，统一由 HasWritePreview 判定
            Execute.Enabled = idle && HasWritePreview();
            Stop.Enabled = busy || canUndo;
            Confirm.Enabled = !busy && canUndo;
            CopyLog.Enabled = CloseButton.Enabled = true;
        }

        /// <summary>执行成功后由 Connect 调用：进入可还原状态。</summary>
        internal void MarkWritten()
        { canUndo = true; RefreshActions(); }

        public Size MeasureDefaultClient()
        {
            root.PerformLayout();
            return new Size(DialogSizing.LogicalClient(this).Width, 640);
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
                if(CancelPreview!=null){CancelPreview();base.OnFormClosing(e);return;}
                e.Cancel = true;
                AppendLog("正在处理，请等待当前步骤结束再关闭。");
                return;
            }
            if (canUndo)
                AppendLog("本批仍可还原；如需保留现场，请关闭前先点「取消操作」或「确认操作」。");
            base.OnFormClosing(e);
        }

        /// <summary>
        /// 初始最大化由 DialogSizing 在 Load 完成、首次显示之前统一应用。
        /// Shown 只定位分隔条，不再改变窗口尺寸；预览保持正常窗口状态。
        /// </summary>
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (Ui.IsInertPreview(this) || DialogPreviewFactory.IsPreview(this) || DialogPreviewFactory.IsConstructing)
            {
                // 预览不最大化，但仍要让分隔条居中（宽度此时已确定）
                CenterMapSplitter();
                return;
            }
            // 最大化后宽度才是最终值，再居中一次
            CenterMapSplitter();
        }

        // 非模态显示，与链接助手/按列筛选/表头匹配合并等其它弹窗完全一致。

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            // 刷出句柄创建前累积的日志（如打开窗体时读取工作表失败的提示）
            if (pendingLog.Length > 0)
            {
                logView.AppendText(pendingLog.ToString());
                pendingLog.Length = 0;
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // 来源区/目标区（行 0/3）不恢复行高：它们贴内容 AutoSize，恢复只会带回留白
            RestoreRowHeight(5, 7, "row.map.v2");
            RestoreRowHeight(7, 9, "row.entries.v2");
            RestoreRowHeight(9, 11, "row.report.v2", true);
            CenterMapSplitter();
            // 自愈兜底：打开窗体时（Show 之前）的 SheetsRequested 若因 COM 忙碌等原因
            // 静默失败，来源/目标两个下拉会是空的（Text 由 SetSource 直接赋值所以看不出来）。
            // 显示后再补读一次；RunLoadEliminationSheets 只读表名，可安全重试。
            if (sourcePath.Text != null && sourcePath.Text.Trim().Length > 0
                && sourceSheet.Items.Count == 0 && SheetsRequested != null)
            {
                AppendLog("来源工作表下拉为空，正在重新读取…");
                SheetsRequested(sourcePath.Text.Trim());
            }
        }

        private void RestoreRowHeight(int above, int below, string key)
        { RestoreRowHeight(above, below, key, false); }

        private void RestoreRowHeight(int above, int below, string key, bool fixedBelow)
        {
            int saved = DialogFormatting.LoadRowHeight(this, key);
            if (saved <= 0) return;
            saved = Ui.Px(this, saved);
            int[] heights = root.GetRowHeights();
            if (above >= heights.Length || below >= heights.Length) return;
            if (fixedBelow)
            {
                if (saved < Ui.Px(this, 80)) return;
                root.RowStyles[above].SizeType = SizeType.Absolute; root.RowStyles[above].Height = saved;
                root.PerformLayout();
                return;
            }
            int belowHeight = heights[below] + (heights[above] - saved);
            if (belowHeight < Ui.Px(this, 80)) return;
            root.RowStyles[above].SizeType = SizeType.Absolute; root.RowStyles[above].Height = saved;
            root.RowStyles[below].SizeType = SizeType.Absolute; root.RowStyles[below].Height = belowHeight;
            root.PerformLayout();
        }
    }
}
