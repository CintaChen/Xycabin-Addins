using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

namespace Excel_Iterate
{
    /// <summary>追踪模式：引用（前向）或从属（反向）。</summary>
    public enum TrackerMode { Precedents, Dependents }

    // =====================================================================
    // TrackerForm — 「追踪引用/从属」非模态弹窗
    //
    // Precedents 模式：单表平铺引用项（同步扫描，快）。
    // Dependents 模式：后台线程遍历所有已打开工作簿反向匹配，
    //   带进度条 + 取消按钮。
    // 复用参考项目 LinkViewForm 的非模态窗/标黄 Timer 模式。
    // =====================================================================

    internal class TrackerForm : Form
    {
        private readonly object _app;
        private readonly TrackerMode _mode;

        private PrecedentResult _precResult;
        private DependentResult _depResult;

        private TableLayoutPanel _root;
        private Label _lblSummary;
        private Label _lblEmpty;
        private DataGridView _grid;
        private RoundedButton _btnJump;
        private RoundedButton _btnErrorSource;
        private RoundedButton _btnCopy;
        private RoundedButton _btnRefresh;
        private RoundedButton _btnClose;
        private RoundedButton _btnCancel;
        private ProgressBar _progress;
        private ContextMenuStrip _menu;
        private Label _lblTip;
        private Timer _tipTimer;

        // 标黄三件套
        private object _flashRange;
        private int _flashPrev;
        private Timer _flashTimer;

        // 「显示错误来源」：行错误状态缓存（行号 → 该行单元格是否错误）与结果窗
        private readonly Dictionary<int, bool> _errorFlags = new Dictionary<int, bool>();
        private ErrorSourceForm _errorForm;

        // 后台扫描（Dependents）
        private HostBatchRunner<FormulaBlock,DependentResult> _scanSession;
        private volatile bool _cancelRequested;

        /// <summary>当前窗口的追踪模式（供 Connect 判断已开窗口是否同模式）。</summary>
        public TrackerMode Mode { get { return _mode; } }

        public TrackerForm(object application, TrackerMode mode)
        {
            _app = application;
            _mode = mode;
            bool preview=DialogPreviewFactory.IsConstructing;
            BuildUi();
            Load += delegate
            {
                if(preview)return;
                // Load 在 Show() 调用栈内、功能区回调线程上同步触发：若在这里直接跑
                // COM 解析/全簿扫描，回调不返回会淹没宿主消息泵（表现为宿主卡顿）。
                // BeginInvoke 把业务排到消息泵 idle 后执行，Load 立即返回，
                // 与 CloseTrackerForm 的 BeginInvoke 推迟模式一致。
                BeginInvoke((MethodInvoker)delegate
                {
                    if (IsDisposed) return;
                    if (_mode == TrackerMode.Precedents) RefreshPrecedents();
                    else StartDependentsScan();
                });
            };
            FormClosed += delegate
            {
                _cancelRequested = true;
                if(_scanSession!=null){try{_scanSession.Dispose();}catch(Exception error){UiDiagnostics.Record("scan-close",error);}_scanSession=null;}
                if (_tipTimer != null) _tipTimer.Stop();
                if(!preview)ClearPendingFlash();
                // 兜底：万一关闭路径上有异常把宿主根窗口留在禁用态
                // （表现为"关窗后工作表/功能区锁死"），这里恢复。
                // 内部 workers!=0 时自动跳过，不会破坏正常 modal 等待语义。
                try { HostDialogService.EnsureHostEnabled(); } catch { }
            };
        }

        // ---------------- 界面构建 ----------------

        private void BuildUi()
        {
            bool green = _mode == TrackerMode.Dependents;
            Text = green ? "追踪从属" : "追踪引用";
            DialogSizing.Prepare(this, green ? "trace.dependents" : "trace.precedents",
                green ? "dependents" : "precedents", green ? Ui.Green : Ui.Blue);
            // 窗口样式与链接助手/按列筛选/合并抵消完全对齐：不设 ShowInTaskbar、
            // MinimizeBox——tool window 样式（WS_EX_TOOLWINDOW）在 WPS 宿主线程会
            // 抢焦点/失激活，表现为宿主单元格无法输入、功能区子菜单点不开
            // （与 MergeHeadersForm 同类结论）。此前误判为 TopMost/加载期同步业务，
            // 均非根因；样式对齐后弹窗仍为非模态、功能不变。
            MaximizeBox = true;
            AcceptButton = null; CancelButton = null; KeyPreview = true;
            KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Escape) { e.Handled = true; Close(); }
                if (e.Control && e.Shift && e.KeyCode == Keys.C) { e.Handled = true; CopySelectedFullPath(); }
                if (e.Control && e.Shift && e.KeyCode == Keys.F) { e.Handled = true; CopySelectedFormula(); }
            };
            _root = new TableLayoutPanel();
            _root.Dock = DockStyle.Fill; _root.Padding = new Padding(16);
            _root.ColumnCount = 1; _root.RowCount = 5;
            _root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
            _root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, green ? 28 : 0));
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            UiCard header = new UiCard(); header.Dock = DockStyle.Fill;
            TableLayoutPanel heading = new TableLayoutPanel(); heading.Dock = DockStyle.Fill;
            heading.ColumnCount = 1; heading.RowCount = 3;
            heading.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            heading.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            heading.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            heading.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Label title = Ui.Label(green ? "追踪从属" : "追踪引用", 16f, true);
            title.ForeColor = green ? Ui.Green : Ui.Blue; heading.Controls.Add(title, 0, 0);
            _lblSummary = Ui.Label(green ? "正在扫描所有已打开工作簿…" : "正在解析…", 10f, false);
            ToolTip summaryTip = new ToolTip();
            _lblSummary.TextChanged += delegate { summaryTip.SetToolTip(_lblSummary, _lblSummary.Text); UpdateEmptyState(); };
            Disposed += delegate { summaryTip.Dispose(); };
            heading.Controls.Add(_lblSummary, 0, 1);
            Label hint = Ui.Label("双击记录可跳转并临时标黄 · 右键可复制更多信息", 9f, false);
            hint.ForeColor = Ui.InkDim; heading.Controls.Add(hint, 0, 2);
            header.Controls.Add(heading); _root.Controls.Add(header, 0, 0);
            // 结果表格
            _grid = new DataGridView();
            _grid.Dock = DockStyle.Fill;
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.AllowUserToResizeRows = false;
            _grid.RowHeadersVisible = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _grid.BackgroundColor = Color.White;
            _grid.BorderStyle = BorderStyle.None;            // 外框交给下面 1px gridHost
            _grid.EnableHeadersVisualStyles = false;         // ★ 关键：否则自定义表头色不生效
            _grid.ColumnHeadersHeightSizeMode =
                DataGridViewColumnHeadersHeightSizeMode.DisableResizing;   // 先锁，再设高度
            _grid.ColumnHeadersHeight = 30;
            _grid.ColumnHeadersDefaultCellStyle.BackColor = Ui.GridHead;
            _grid.ColumnHeadersDefaultCellStyle.ForeColor = Ui.GridHeadInk;
            _grid.ColumnHeadersDefaultCellStyle.Font = new Font(Font.FontFamily, 9f, FontStyle.Bold);
            _grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            _grid.RowTemplate.Height = 24;                   // ★ 必须在 Rows.Add 之前设
            _grid.DefaultCellStyle.BackColor = Color.White;
            _grid.DefaultCellStyle.ForeColor = Ui.Ink;
            _grid.DefaultCellStyle.Font = new Font(Font.FontFamily, 9f);
            _grid.DefaultCellStyle.Padding = new Padding(4, 0, 4, 0);
            _grid.DefaultCellStyle.SelectionBackColor = Ui.SelectBack;
            _grid.DefaultCellStyle.SelectionForeColor = Ui.SelectInk;
            _grid.AlternatingRowsDefaultCellStyle.BackColor = Ui.ZebraBack;
            _grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;  // 只留横线
            _grid.GridColor = Ui.GridLine;
            _grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            _grid.CellDoubleClick += delegate { JumpToSelectedItem(); };
            _grid.CellMouseDown += delegate(object s, DataGridViewCellMouseEventArgs e)
            {
                if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    _grid.ClearSelection();
                    _grid.Rows[e.RowIndex].Selected = true;
                }
            };
            _grid.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter) { e.Handled = true; JumpToSelectedItem(); }
            };
            _grid.SelectionChanged += delegate { UpdateErrorSourceButton(); };
            // 长公式在窄列里必然显示不全：悬停弹出全文（自动折行，超长截断）。
            // 只在「确实被截断」时才给 ToolTip，短文本悬停不会打扰用户。
            _grid.CellToolTipTextNeeded += delegate(object s, DataGridViewCellToolTipTextNeededEventArgs e)
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                DataGridViewCell cell = _grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                string text = Convert.ToString(cell.Value);
                if (string.IsNullOrEmpty(text)) return;

                // 注意：这里不能带 EndEllipsis 量（那会把文本裁到刚好放下），
                // 要量「完整文本所需宽度」再和实际可用宽度比。
                Size avail = new Size(cell.Size.Width - cell.InheritedStyle.Padding.Horizontal - 2, cell.Size.Height);
                Size need = TextRenderer.MeasureText(text, cell.InheritedStyle.Font ?? _grid.Font);
                if (need.Width <= avail.Width) return;       // 没被截断就不弹

                e.ToolTipText = WrapForTip(text, 110);
            };

            if (_mode == TrackerMode.Precedents)
            {
                MakeCol("#", 40);
                MakeCol("类型", 96);          // 5 个汉字（所选/同表/跨表/跨簿 + 单元格）要完整显示
                MakeCol("工作簿", 150, 260);  // 长名封顶，自动列宽不撑爆
                MakeCol("工作表", 140, 220);
                MakeCol("单元格", 120);
                MakeCol("绝对/相对引用", 112);   // 表头 6 汉字 + 斜杠，加宽到一次完整显示
                MakeCol("公式片段", 340, 400);   // 自身行放整条公式，设上限
                MakeCol("完整路径", 460);     // 最后一列，下面设为 Fill 补满窗口
            }
            else
            {
                MakeCol("#", 40);
                MakeCol("工作簿", 180, 260);
                MakeCol("工作表", 150, 220);
                MakeCol("单元格", 120);
                MakeCol("命中片段", 180, 260);
                MakeCol("公式原文", 400, 450);
                MakeCol("完整路径", 300);
            }
            // 最后一列「完整路径」设为 Fill：短列贴合后自动补满窗口，无右侧空白
            _grid.Columns[_grid.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Ui.StyleGrid(_grid);
            if (green)
            {
                _grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(229, 245, 239);
                _grid.DefaultCellStyle.SelectionForeColor = Ui.GreenDown;
            }
            Ui.AutoFitColumns(_grid);
            _grid.Columns[_mode == TrackerMode.Precedents ? 6 : 5].DefaultCellStyle.Font = Ui.MonoFont(10f, Font);
            BuildContextMenu();


            Panel body = new Panel(); body.BackColor = Color.White;
            _lblEmpty = Ui.Label("正在准备追踪结果…", 11f, false);
            _lblEmpty.TextAlign = ContentAlignment.MiddleCenter; _lblEmpty.ForeColor = Ui.InkDim;
            body.Controls.Add(_grid); body.Controls.Add(_lblEmpty);
            _grid.RowsAdded += delegate { UpdateEmptyState(); };
            _grid.RowsRemoved += delegate { UpdateEmptyState(); };
            _root.Controls.Add(Ui.Card(body, "追踪结果"), 0, 1);
            if (green)
            {
                _progress = new UiProgressBar(); _progress.Dock = DockStyle.Fill;
                _progress.Style = ProgressBarStyle.Continuous; _progress.Minimum = 0; _progress.Maximum = 100;
                _progress.Margin = new Padding(0, 4, 0, 4); _root.Controls.Add(_progress, 0, 2);
            }
            _lblTip = Ui.Label("", 9f, false); _lblTip.ForeColor = Ui.Green;
            _root.Controls.Add(_lblTip, 0, 3);
            FlowLayoutPanel bar = Ui.ActionBar();
            _btnJump = Ui.Button("跳转", 96); Ui.StylePrimary(_btnJump, green);
            _btnJump.Click += delegate { JumpToSelectedItem(); }; bar.Controls.Add(_btnJump);
            if (!green)
            {
                // 「显示错误来源」：选中错误行（或起点自身行）时可用，
                // 点击后沿引用链回溯，列出导致错误的全部源头格。
                _btnErrorSource = Ui.Button("显示错误来源", 130);
                _btnErrorSource.Enabled = false;
                _btnErrorSource.Click += delegate { ShowErrorSources(); }; bar.Controls.Add(_btnErrorSource);
                ToolTip tipErr = new ToolTip();
                tipErr.SetToolTip(_btnErrorSource, "沿引用链回溯，列出导致所选单元格报错的源头单元格");
                Disposed += delegate { tipErr.Dispose(); };
            }
            _btnCopy = Ui.Button("复制完整路径", 140);
            _btnCopy.Click += delegate { CopySelectedFullPath(); }; bar.Controls.Add(_btnCopy);
            ToolTip tip = new ToolTip(); tip.SetToolTip(_btnCopy, "复制所选行的工作簿路径（Ctrl+Shift+C）\r\n右键列表可复制引用全址、公式和选中行");
            Disposed += delegate { tip.Dispose(); };
            if (!green)
            { _btnRefresh = Ui.Button("刷新", 96); _btnRefresh.Click += delegate { RefreshPrecedents(); }; bar.Controls.Add(_btnRefresh); }
            else
            { _btnCancel = Ui.Button("取消扫描", 110); _btnCancel.Click += delegate { _cancelRequested = true; if(_scanSession!=null)_scanSession.Cancel(); }; bar.Controls.Add(_btnCancel); }
            _btnClose = Ui.Button("关闭", 96); _btnClose.Click += delegate { Close(); }; bar.Controls.Add(_btnClose);
            _root.Controls.Add(bar, 0, 4); Controls.Add(_root);
            UpdateEmptyState();
        }

        private void UpdateEmptyState()
        {
            if (_lblEmpty == null || _grid == null) return;
            if (_btnJump != null) _btnJump.Enabled = _grid.Rows.Count > 0;
            if (_btnCopy != null) _btnCopy.Enabled = _grid.Rows.Count > 0;
            UpdateErrorSourceButton();
            _lblEmpty.Visible = _grid.Rows.Count == 0;
            _lblEmpty.Text = _lblSummary.Text + "\r\n\r\n刷新或重新选择单元格后，可再次追踪。";
            if (_grid.Rows.Count == 0) _lblEmpty.BringToFront();
        }

        // ---------------- 「显示错误来源」 ----------------

        /// <summary>
        /// 判定某行对应单元格是否为错误格（值错误或公式含错误字面量），带行级缓存。
        /// 文本层先判（零 COM 开销）；未命中再经 ErrorSourceEngine.ReadCell 读值，
        /// 外部工作簿不自动打开（autoOpen=false），打不开按非错误处理。
        /// </summary>
        private bool IsErrorRow(int rowIndex)
        {
            bool known;
            if (_errorFlags.TryGetValue(rowIndex, out known)) return known;
            known = false;
            if (rowIndex >= 0 && rowIndex < _grid.Rows.Count)
            {
                PrecedentItem it = _grid.Rows[rowIndex].Tag as PrecedentItem;
                if (it != null)
                {
                    known = FormulaRefParser.ContainsErrorLiteral(it.FormulaFragment);
                    if (!known && _app != null && !DialogPreviewFactory.IsConstructing)
                    {
                        try
                        {
                            CellReadResult cr = ErrorSourceEngine.ReadCell(_app,
                                it.WorkbookPath, it.WorkbookName, it.SheetName, it.CellAddressA1, false);
                            known = cr.Found && cr.HasErrorValue;
                        }
                        catch { }
                    }
                }
            }
            _errorFlags[rowIndex] = known;
            return known;
        }

        /// <summary>刷新后把所有错误行标红（IsErrorRow 的缓存顺带填满，供按钮判定复用）。</summary>
        private void MarkErrorRows()
        {
            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                if (!IsErrorRow(i)) continue;
                _grid.Rows[i].DefaultCellStyle.ForeColor = Ui.Danger;
                // 第 0 行自身行有独立选中配色，选中前景色保持不动，只改未选中前景
            }
        }

        /// <summary>「显示错误来源」仅对错误行（含第 0 行起点自身）可用。</summary>
        private void UpdateErrorSourceButton()
        {
            if (_btnErrorSource == null) return;
            bool on = false;
            if (_mode == TrackerMode.Precedents && _grid.SelectedRows.Count > 0)
                on = IsErrorRow(_grid.SelectedRows[0].Index);
            _btnErrorSource.Enabled = on;
        }

        /// <summary>对选中行启动错误来源回溯，结果用 ErrorSourceForm 展示。</summary>
        private void ShowErrorSources()
        {
            if (DialogPreviewFactory.IsPreview(this)) return;
            if (_grid.SelectedRows.Count == 0) return;
            PrecedentItem it = _grid.SelectedRows[0].Tag as PrecedentItem;
            if (it == null) return;

            Cursor.Current = Cursors.WaitCursor;
            ErrorTraceResult result;
            try
            {
                result = ErrorSourceEngine.TraceFrom(_app,
                    it.WorkbookPath, it.WorkbookName, it.SheetName,
                    new List<string> { it.CellAddressA1 }, null, null);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

            if (!result.Success)
            {
                UiDialog.Show(this, "错误来源追溯失败：" + result.ErrorMessage,
                    "显示错误来源", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (result.Sources.Count == 0)
            {
                ShowTip("未发现错误来源：该单元格不是错误值，或引用链上均无错误。");
                return;
            }

            if (_errorForm != null && !_errorForm.IsDisposed) _errorForm.Close();
            _errorForm = new ErrorSourceForm(_app, result, it.CellAddressA1);
            _errorForm.Show(this);   // owned 窗：本窗关闭时联动关闭
        }

        private void MakeCol(string header, int width, int maxWidth = 0)
        {
            DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
            c.HeaderText = header;
            c.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            c.MinimumWidth = 50;
            c.SortMode = DataGridViewColumnSortMode.NotSortable;

            // 表头按「实际渲染字体」量宽（TextRenderer 是 DPI 感知的），
            // 保证任何 DPI 缩放 / 字体回退下都完整显示，宽度只增不减。
            Font hf = _grid.ColumnHeadersDefaultCellStyle.Font ?? _grid.Font;
            int need = TextRenderer.MeasureText(header, hf).Width + 32;   // 24 = 左右内边距 + 分隔线 + 余量
            c.Width = Math.Max(width, need);

            // 可选上限（长文本列用，自动列宽后手动封顶）。DataGridViewColumn 没有 MaximumWidth，
            // 上限暂存在列 Tag（列 Tag 本项目未作他用），AutoFitColumns 里贴合后再按此封顶。
            if (maxWidth > 0) c.Tag = Math.Max(maxWidth, need);

            _grid.Columns.Add(c);
        }

        /// <summary>
        /// 数据填充后，对每列做「贴合最宽内容」（等价双击列宽线 = AutoResize(AllCells)）。
        /// 跳过最后一列（Fill 列自行补满窗口）。长文本列贴合后再按列 Tag 里的上限手动封顶
        /// （DataGridViewColumn 无 MaximumWidth，故用 Tag 暂存上限）。
        /// </summary>
        private void AutoFitColumns()
        { Ui.AutoFitColumns(_grid); }

        private void BuildContextMenu()
        {
            _menu = new ContextMenuStrip();
            _menu.Font = Ui.MakeFont(10f, FontStyle.Regular, Font);
            _menu.Renderer = new ToolStripProfessionalRenderer();
            AddMenuItem(_menu, "复制完整路径（工作簿文件）", delegate { CopySelectedFullPath(); });
            AddMenuItem(_menu, "复制引用全址（路径+表+单元格）", delegate { CopySelectedFullRef(); });
            AddMenuItem(_menu, "复制单元格地址", delegate { CopySelectedCellAddress(); });
            AddMenuItem(_menu, "复制公式（Ctrl+Shift+F）", delegate { CopySelectedFormula(); });
            _menu.Items.Add(new ToolStripSeparator());
            AddMenuItem(_menu, "复制选中行", delegate { CopyRows(false); });
            AddMenuItem(_menu, "复制全部行", delegate { CopyRows(true); });
            _grid.ContextMenuStrip = _menu;
        }

        private static void AddMenuItem(ContextMenuStrip menu, string text, EventHandler onClick)
        {
            ToolStripMenuItem mi = new ToolStripMenuItem(text);
            mi.Click += onClick;
            menu.Items.Add(mi);
        }

        // ---------------- Precedents：同步解析 ----------------

        private void RefreshPrecedents()
        {
            _lblSummary.Text = "正在解析…";
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                _precResult = PrecedentTrackerEngine.TracePrecedents(_app);
                _grid.Rows.Clear();
                _errorFlags.Clear();

                if (_precResult.Status != TrackerStatus.Completed)
                {
                    _lblSummary.Text = StatusToText(_precResult.Status);
                    return;                       // 无公式/无簿/无活动格：不加「所选单元格」行
                }

                _lblSummary.Text = string.Format(
                    "当前单元格：{0} | {1} | {2}  —  引用 {3} 处",
                    _precResult.WorkbookName, _precResult.SheetName,
                    _precResult.CellAddress, _precResult.Items.Count);

                AddSelfRow(_precResult.Formula);   // 第 0 行：所选单元格（起点自身，含整条公式）

                foreach (PrecedentItem item in _precResult.Items)
                {
                    int r = _grid.Rows.Add(
                        item.Index,
                        TypeText(item.RefType),
                        item.WorkbookName,
                        item.SheetName,
                        AddrText(item),
                        AbsText(item),
                        OneLine(item.FormulaFragment),
                        item.WorkbookPath);
                    _grid.Rows[r].Tag = item;
                }
                if (_grid.Rows.Count > 0) _grid.Rows[0].Selected = true;
                MarkErrorRows();           // 错误行红字（顺带填充 _errorFlags 缓存）
                AutoFitColumns();
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        // ---------------- 「所选单元格」自身行 ----------------

        /// <summary>
        /// 第 0 行「所选单元格」：追踪的起点自身，浅黄底加粗，双击可跳回并标黄。
        /// 只存在于 UI 层，不进 _precResult.Items（保住 Items.Count 的「引用 N 处」语义）。
        /// Tag 是标准 PrecedentItem，跳转/复制全址等能力天然复用，无需改别处。
        /// 第七列填自己的整条公式（其余引用行填的是「命中的那一片段」，语义不同但同列）。
        /// </summary>
        private void AddSelfRow(string formula)
        {
            PrecedentItem self = new PrecedentItem();
            self.Index = 0;
            self.RefType = RefType.SameSheet;
            self.WorkbookName = _precResult.WorkbookName;
            self.WorkbookPath = _precResult.WorkbookPath;
            self.SheetName = _precResult.SheetName;
            self.CellAddressA1 = _precResult.CellAddress;
            self.RangeEndA1 = "";
            self.FormulaFragment = formula ?? "";   // 保留原始文本（含前置 =），供复制/悬停
            self.IsRange = false;
            self.IsColAbsolute = false;
            self.IsRowAbsolute = false;

            int r = _grid.Rows.Add(
                0,                       // # 序号：0 = 起点，后面引用项保持 1..N
                "所选单元格",             // 类型（不走 TypeText）
                self.WorkbookName,
                self.SheetName,
                self.CellAddressA1,
                "—",                     // 绝对/相对：自身不适用
                OneLine(self.FormulaFragment),   // 网格是单行显示，换行需归一
                self.WorkbookPath);

            DataGridViewRow row = _grid.Rows[r];
            row.Tag = self;
            // 行级样式优先级高于斑马纹，选中色也要单独指定，否则选中后看不出是自身行
            row.DefaultCellStyle.BackColor = Ui.SelfRowBack;
            row.DefaultCellStyle.SelectionBackColor = Ui.SelfRowSel;
            row.DefaultCellStyle.SelectionForeColor = Color.FromArgb(30, 30, 30);
            row.DefaultCellStyle.Font = new Font(_grid.Font, FontStyle.Bold);
            row.Cells[6].Style.Font = Ui.MonoFont(10f, _grid.Font);
        }

        // ---------------- Dependents：后台扫描 ----------------

        private void StartDependentsScan()
        {
            _cancelRequested = false;
            if (_progress != null) { _progress.Visible = true; _progress.Value = 0; }
            if (_btnCancel != null) _btnCancel.Enabled = true;
            _lblSummary.Text = "正在扫描所有已打开工作簿…";

            if(_scanSession!=null)_scanSession.Dispose();
            var result = new DependentResult();
            try
            {
                var reader = new DependentSnapshotReader(_app,result);
                if(!reader.Valid){reader.Dispose();FillDependents(result);return;}
                if(!reader.CanYield)_lblSummary.Text="宿主不支持来源变化监听，正在同步读取；随后后台解析…";
                var accumulator = new DependentAccumulator(result);
                HostBatchRunner<FormulaBlock,DependentResult> session=null;
                session=new HostBatchRunner<FormulaBlock,DependentResult>(reader,accumulator.Accept,accumulator.Finish,
                    delegate(DependentResult found,Exception error,bool cancelled)
                    {
                        if(IsDisposed||!ReferenceEquals(_scanSession,session))return;
                        _scanSession=null;found=found??result;
                        if(error!=null){found.Status=TrackerStatus.Error;found.ErrorMessage=error.Message;}
                        found.Cancelled=cancelled;FillDependents(found);
                    },delegate(FormulaBlock block)
                    {
                        if(IsDisposed)return;
                        if(_progress!=null&&block.WorkbookCount>0)_progress.Value=Math.Max(0,Math.Min(100,(int)(100.0*(block.WorkbookIndex-1+(double)block.SheetIndex/block.SheetCount)/block.WorkbookCount)));
                        _lblSummary.Text="正在扫描："+block.SheetName;
                    });
                _scanSession=session;_btnRefresh.Enabled=false;session.Start();
            }
            catch(Exception error){result.Status=TrackerStatus.Error;result.ErrorMessage=error.Message;FillDependents(result);}
        }

        private void FillDependents(DependentResult result)
        {
            _depResult = result;
            _grid.Rows.Clear();
            DoneScan();

            if (result.Cancelled)
            {
                _lblSummary.Text = "已取消扫描（已找到 " + result.Items.Count + " 处）。";
            }
            else if (result.Status != TrackerStatus.Completed)
            {
                _lblSummary.Text = result.Status == TrackerStatus.Error
                    ? "扫描出错：" + (result.ErrorMessage ?? "未知")
                    : StatusToText(result.Status);
            }
            else _lblSummary.Text = string.Format(
                "目标单元格：{0} | {1} | {2}  —  从属 {3} 处（扫描 {4} 格 / {5} 个公式）",
                result.TargetWorkbookName, result.TargetSheetName,
                result.TargetCellAddress, result.Items.Count,
                result.ScannedCells, result.ScannedFormulas);

            var batch=new List<DataGridViewRow>();
            foreach(DependentItem item in result.Items)
            {
                var row=new DataGridViewRow();row.CreateCells(_grid,item.Index,item.WorkbookName,item.SheetName,item.CellAddressA1,OneLine(item.MatchFragment),OneLine(item.SourceFormula),item.WorkbookPath);row.Tag=item;batch.Add(row);
            }
            _grid.SuspendLayout();try{_grid.Rows.AddRange(batch.ToArray());}finally{_grid.ResumeLayout(false);}
            if (_grid.Rows.Count > 0) _grid.Rows[0].Selected = true;
            AutoFitColumns();
        }

        private void DoneScan()
        {
            if (_progress != null) _progress.Visible = false;
            if (_btnCancel != null) _btnCancel.Enabled = false;
            if (_btnRefresh != null) _btnRefresh.Enabled = true;
        }

        // ---------------- 复制 ----------------

        /// <summary>取当前选中行绑定的数据项（PrecedentItem 或 DependentItem）。</summary>
        private object SelectedItem()
        {
            if (_grid.SelectedRows.Count == 0) return null;
            return _grid.SelectedRows[0].Tag;
        }

        /// <summary>选中项的完整工作簿路径（即「完整路径」列内容）。</summary>
        private static string FullPathOf(object tag)
        {
            PrecedentItem p = tag as PrecedentItem;
            if (p != null) return p.WorkbookPath;
            DependentItem d = tag as DependentItem;
            if (d != null) return d.WorkbookPath;
            return "";
        }

        /// <summary>选中项的「引用全址」：路径 + 工作簿 + 工作表 + 单元格，可直接粘进公式。</summary>
        private static string FullRefOf(object tag)
        {
            PrecedentItem p = tag as PrecedentItem;
            if (p != null)
            {
                string addr = AddrText(p);
                if (p.RefType == RefType.CrossWorkbook && !string.IsNullOrEmpty(p.WorkbookPath))
                {
                    string dir = p.WorkbookPath.TrimEnd('\\');
                    return "'" + dir + "\\[" + p.WorkbookName + "]" + p.SheetName + "'!" + addr;
                }
                string head = string.IsNullOrEmpty(p.WorkbookName)
                    ? p.SheetName
                    : "[" + p.WorkbookName + "]" + p.SheetName;
                return head + "!" + addr;
            }
            DependentItem d = tag as DependentItem;
            if (d != null)
            {
                string dir = (d.WorkbookPath ?? "").TrimEnd('\\');
                if (dir.Length > 0)
                    return "'" + dir + "\\[" + d.WorkbookName + "]" + d.SheetName + "'!" + d.CellAddressA1;
                return "[" + d.WorkbookName + "]" + d.SheetName + "!" + d.CellAddressA1;
            }
            return "";
        }

        /// <summary>选中项的单元格地址（区域显示为首格:末格）。</summary>
        private static string CellAddressOf(object tag)
        {
            PrecedentItem p = tag as PrecedentItem;
            if (p != null) return AddrText(p);
            DependentItem d = tag as DependentItem;
            if (d != null) return d.CellAddressA1;
            return "";
        }

        private void CopySelectedFullPath()
        {
            object tag = SelectedItem();
            if (tag == null) { ShowTip("请先在列表里选中一行。"); return; }
            string path = FullPathOf(tag);
            if (string.IsNullOrEmpty(path)) { ShowTip("该行没有完整路径。"); return; }
            CopyToClipboard(path, "已复制完整路径：", "剪贴板写入失败，请重试。");
        }

        private void CopySelectedFullRef()
        {
            object tag = SelectedItem();
            if (tag == null) { ShowTip("请先在列表里选中一行。"); return; }
            string s = FullRefOf(tag);
            if (string.IsNullOrEmpty(s)) { ShowTip("该行无法生成引用全址。"); return; }
            CopyToClipboard(s, "已复制引用全址：", "剪贴板写入失败，请重试。");
        }

        private void CopySelectedCellAddress()
        {
            object tag = SelectedItem();
            if (tag == null) { ShowTip("请先在列表里选中一行。"); return; }
            string s = CellAddressOf(tag);
            if (string.IsNullOrEmpty(s)) { ShowTip("该行没有单元格地址。"); return; }
            CopyToClipboard(s, "已复制单元格地址：", "剪贴板写入失败，请重试。");
        }

        /// <summary>
        /// 选中行的公式：引用窗 = 自身行的整条公式 / 引用行的命中片段；
        /// 从属窗 = 该单元格的公式原文。复制的是原文（保留换行与前置 =）。
        /// </summary>
        private void CopySelectedFormula()
        {
            object tag = SelectedItem();
            if (tag == null) { ShowTip("请先在列表里选中一行。"); return; }
            string s = FormulaOf(tag);
            if (string.IsNullOrEmpty(s)) { ShowTip("该行没有公式。"); return; }
            CopyToClipboard(s, "已复制公式：", "剪贴板写入失败，请重试。");
        }

        private static string FormulaOf(object tag)
        {
            PrecedentItem p = tag as PrecedentItem;
            if (p != null) return p.FormulaFragment;
            DependentItem d = tag as DependentItem;
            if (d != null) return d.SourceFormula;
            return "";
        }

        /// <summary>把表格内容复制成 TSV（带表头），可直接粘进 Excel / WPS 表格。</summary>
        private void CopyRows(bool all)
        {
            if (_grid.Rows.Count == 0) { ShowTip("列表为空，没有可复制的内容。"); return; }

            StringBuilder sb = new StringBuilder();
            for (int c = 0; c < _grid.Columns.Count; c++)
            {
                if (c > 0) sb.Append('\t');
                sb.Append(CellToTsv(_grid.Columns[c].HeaderText));
            }
            int copied = 0;
            for (int r = 0; r < _grid.Rows.Count; r++)
            {
                DataGridViewRow row = _grid.Rows[r];
                if (!all && !row.Selected) continue;
                sb.Append("\r\n");
                for (int c = 0; c < _grid.Columns.Count; c++)
                {
                    if (c > 0) sb.Append('\t');
                    object v = row.Cells[c].Value;
                    sb.Append(CellToTsv(v == null ? "" : Convert.ToString(v)));
                }
                copied++;
            }
            if (copied == 0) { ShowTip("请先在列表里选中一行。"); return; }
            CopyToClipboard(sb.ToString(),
                all ? "已复制全部 " + copied + " 行：" : "已复制选中行：",
                "剪贴板写入失败，请重试。");
        }

        private static string CellToTsv(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
        }

        /// <summary>写剪贴板（保留一份，宿主关闭后仍可粘贴）；失败重试 3 次。</summary>
        private static bool SetClipboard(string text)
        {
            for (int i = 0; i < 3; i++)
            {
                try { Clipboard.SetDataObject(text, true); return true; }
                catch { Thread.Sleep(80); }
            }
            return false;
        }

        private void CopyToClipboard(string text, string okPrefix, string failMsg)
        {
            if (SetClipboard(text)) ShowTip(okPrefix + TrimForTip(text));
            else ShowTip(failMsg);
        }

        /// <summary>长文本在提示条里做中间省略。</summary>
        private static string TrimForTip(string s)
        {
            const int Max = 46;
            if (string.IsNullOrEmpty(s) || s.Length <= Max) return s ?? "";
            return s.Substring(0, 28) + "…" + s.Substring(s.Length - 15);
        }

        /// <summary>在按钮栏右侧显示一次短暂提示，2.6 秒后自动清空。</summary>
        private void ShowTip(string text)
        {
            if (_lblTip == null) return;
            _lblTip.Text = text;
            if (_tipTimer == null)
            {
                _tipTimer = new Timer();
                _tipTimer.Interval = 2600;
                _tipTimer.Tick += delegate { _tipTimer.Stop(); _lblTip.Text = ""; };
            }
            _tipTimer.Stop();
            _tipTimer.Start();
        }

        // ---------------- 跳转 ----------------

        private void JumpToSelectedItem()
        {
            if(DialogPreviewFactory.IsPreview(this))return;
            if (_grid.SelectedRows.Count == 0) return;
            object tag = _grid.SelectedRows[0].Tag;
            if (tag == null) return;

            object flash;
            int prev;
            string err;
            if (_mode == TrackerMode.Precedents)
            {
                PrecedentItem it = (PrecedentItem)tag;
                err = PrecedentTrackerEngine.JumpToPrecedentEx(_app, it, out flash, out prev);
            }
            else
            {
                DependentItem it = (DependentItem)tag;
                err = PrecedentTrackerEngine.JumpToCellEx(_app,
                    it.WorkbookName, it.WorkbookPath, it.SheetName, it.CellAddressA1,
                    true, out flash, out prev);
            }

            if (err != null)
            {
                UiDialog.Show(this, err, "跳转", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            ClearPendingFlash();
            if (flash != null)
            {
                _flashRange = flash;
                _flashPrev = prev;
                if (_flashTimer == null)
                {
                    _flashTimer = new Timer();
                    _flashTimer.Interval = 2500;
                    _flashTimer.Tick += delegate { ClearPendingFlash(); };
                }
                _flashTimer.Start();
            }
        }

        internal bool CleanupForWorkbook(string path)
        {
            if (_scanSession != null && _scanSession.IsRunning)
            {
                _cancelRequested = true; _scanSession.Cancel();
                throw new InvalidOperationException("从属扫描正在取消；等待扫描结束后重试，本阶段未保存或替换文件。");
            }
            bool related = _precResult != null && string.Equals(_precResult.WorkbookPath,path,StringComparison.OrdinalIgnoreCase);
            related |= _depResult != null && (string.Equals(_depResult.TargetWorkbookName,System.IO.Path.GetFileName(path),StringComparison.OrdinalIgnoreCase) || _depResult.Items.Exists(i=>string.Equals(i.WorkbookPath,path,StringComparison.OrdinalIgnoreCase)));
            if (_flashRange != null) using (ComBag bag = new ComBag())
            {
                dynamic sheet = bag.Own(((dynamic)_flashRange).Worksheet), book = bag.Own(sheet.Parent);
                related |= string.Equals(Convert.ToString(book.FullName),path,StringComparison.OrdinalIgnoreCase);
            }
            if (!related) return false;
            if (_flashRange != null) using (ComBag bag = new ComBag())
            {
                dynamic fill = bag.Own(((dynamic)_flashRange).Interior);
                if (_flashPrev == -4142) { fill.ColorIndex = _flashPrev; if (Convert.ToInt32(fill.ColorIndex) != _flashPrev) throw new InvalidOperationException("临时高亮还原核对失败，未自动保存。"); }
                else { fill.Color = _flashPrev; if (Convert.ToInt32(fill.Color) != _flashPrev) throw new InvalidOperationException("临时高亮还原核对失败，未自动保存。"); }
            }
            ClearPendingFlash(); Close(); return true;
        }
        private void ClearPendingFlash()
        {
            if (_flashTimer != null) _flashTimer.Stop();
            if (_flashRange != null)
            {
                PrecedentTrackerEngine.ClearFlash(_flashRange, _flashPrev);
                _flashRange = null;
            }
        }

        // ---------------- 文案辅助 ----------------

        private static string StatusToText(TrackerStatus status)
        {
            switch (status)
            {
                case TrackerStatus.NoWorkbook: return "未检测到打开的工作簿。";
                case TrackerStatus.NoActiveCell: return "请先选中一个单元格。";
                case TrackerStatus.NoFormula: return "当前单元格没有公式（无法追踪）。";
                case TrackerStatus.Error: return "处理失败。";
                default: return "";
            }
        }

        private static string TypeText(RefType t)
        {
            switch (t)
            {
                case RefType.SameSheet: return "同表单元格";
                case RefType.CrossSheet: return "跨表单元格";
                case RefType.CrossWorkbook: return "跨簿单元格";
                default: return "";
            }
        }

        /// <summary>网格单元格是单行显示，公式里的换行先归一成空格，免得撑爆行高。</summary>
        private static string OneLine(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\r", " ").Replace("\n", " ");
        }

        /// <summary>把长公式按固定字数折行后放进 ToolTip，避免一行拉出屏幕。</summary>
        private static string WrapForTip(string s, int lineChars)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = OneLine(s);
            const int Max = 4000;
            bool cut = s.Length > Max;
            if (cut) s = s.Substring(0, Max);

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < s.Length; i += lineChars)
            {
                if (i > 0) sb.Append("\r\n");
                int n = Math.Min(lineChars, s.Length - i);
                sb.Append(s, i, n);
            }
            if (cut) sb.Append("\r\n…");
            return sb.ToString();
        }

        private static string AddrText(PrecedentItem item)
        {
            if (item.IsRange) return item.CellAddressA1 + ":" + item.RangeEndA1;
            return item.CellAddressA1;
        }

        private static string AbsText(PrecedentItem item)
        {
            if (item.IsColAbsolute && item.IsRowAbsolute) return "$列$行";
            if (item.IsColAbsolute) return "$列";
            if (item.IsRowAbsolute) return "$行";
            return "相对";
        }
    }
}
