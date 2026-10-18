using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

namespace Excel_Iterate
{
    // =====================================================================
    // ErrorSourceForm — 「错误来源」结果弹窗（由追踪引用窗的「显示错误来源」打开）
    //
    // 列出导致所选单元格报错的全部源头格（ErrorSourceEngine.TraceCore 的输出），
    // 支持跳转标黄（复用 PrecedentTrackerEngine.JumpToCellEx）与复制公式/传播路径。
    // 作为 TrackerForm 的 owned 窗打开，主窗关闭时联动关闭。
    // =====================================================================

    internal class ErrorSourceForm : Form
    {
        private readonly object _app;
        private readonly ErrorTraceResult _result;

        private TableLayoutPanel _root;
        private Label _lblSummary;
        private DataGridView _grid;
        private RoundedButton _btnJump;
        private RoundedButton _btnCopyFormula;
        private RoundedButton _btnCopyPath;
        private RoundedButton _btnClose;
        private Label _lblTip;
        private Timer _tipTimer;

        // 标黄三件套（与 TrackerForm 同款：2.5s 后按原值还原）
        private object _flashRange;
        private int _flashPrev;
        private Timer _flashTimer;

        public ErrorSourceForm(object application, ErrorTraceResult result, string rootAddr)
        {
            _app = application;
            _result = result;
            bool preview = DialogPreviewFactory.IsConstructing;
            BuildUi(rootAddr);
            if (!preview && result != null) FillRows();
            FormClosed += delegate
            {
                if (_tipTimer != null) _tipTimer.Stop();
                if (!preview) ClearPendingFlash();
            };
        }

        // ---------------- 界面构建 ----------------

        private void BuildUi(string rootAddr)
        {
            Text = "错误来源";
            DialogSizing.Prepare(this, "trace.errorsources", "errorsources", Ui.Danger);
            ShowInTaskbar = false; MinimizeBox = false; MaximizeBox = true;
            AcceptButton = null; CancelButton = null; KeyPreview = true;
            KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Escape) { e.Handled = true; Close(); }
            };

            _root = new TableLayoutPanel();
            _root.Dock = DockStyle.Fill; _root.Padding = new Padding(16);
            _root.ColumnCount = 1; _root.RowCount = 4;
            _root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 86));
            _root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            UiCard header = new UiCard(); header.Dock = DockStyle.Fill;
            TableLayoutPanel heading = new TableLayoutPanel(); heading.Dock = DockStyle.Fill;
            heading.ColumnCount = 1; heading.RowCount = 2;
            heading.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            heading.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            heading.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Label title = Ui.Label("错误来源", 16f, true);
            title.ForeColor = Ui.Danger; heading.Controls.Add(title, 0, 0);
            int n = _result == null ? 0 : _result.Sources.Count;
            string summary = _result == null
                ? "正在准备错误来源…"
                : "起点 " + (rootAddr ?? "") + " — 共找到 " + n + " 个错误来源";
            if (_result != null && _result.TruncatedByLimit && _result.LimitNote.Length > 0)
                summary += "（" + _result.LimitNote + "，结果可能不完整）";
            _lblSummary = Ui.Label(summary, 10f, false);
            if (_result != null && _result.TruncatedByLimit) _lblSummary.ForeColor = Ui.Danger;
            ToolTip summaryTip = new ToolTip();
            _lblSummary.TextChanged += delegate { summaryTip.SetToolTip(_lblSummary, _lblSummary.Text); };
            Disposed += delegate { summaryTip.Dispose(); };
            heading.Controls.Add(_lblSummary, 0, 1);
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
            _grid.BorderStyle = BorderStyle.None;
            _grid.EnableHeadersVisualStyles = false;
            _grid.ColumnHeadersHeightSizeMode =
                DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            _grid.ColumnHeadersHeight = 30;
            _grid.ColumnHeadersDefaultCellStyle.BackColor = Ui.GridHead;
            _grid.ColumnHeadersDefaultCellStyle.ForeColor = Ui.GridHeadInk;
            _grid.ColumnHeadersDefaultCellStyle.Font = new Font(Font.FontFamily, 9f, FontStyle.Bold);
            _grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            _grid.RowTemplate.Height = 24;
            _grid.DefaultCellStyle.BackColor = Color.White;
            _grid.DefaultCellStyle.ForeColor = Ui.Ink;
            _grid.DefaultCellStyle.Font = new Font(Font.FontFamily, 9f);
            _grid.DefaultCellStyle.Padding = new Padding(4, 0, 4, 0);
            _grid.DefaultCellStyle.SelectionBackColor = Ui.SelectBack;
            _grid.DefaultCellStyle.SelectionForeColor = Ui.SelectInk;
            _grid.AlternatingRowsDefaultCellStyle.BackColor = Ui.ZebraBack;
            _grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            _grid.GridColor = Ui.GridLine;
            _grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            _grid.CellDoubleClick += delegate { JumpToSelected(); };
            _grid.KeyDown += delegate(object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter) { e.Handled = true; JumpToSelected(); }
            };
            _grid.CellToolTipTextNeeded += delegate(object s, DataGridViewCellToolTipTextNeededEventArgs e)
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                DataGridViewCell cell = _grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                string text = Convert.ToString(cell.Value);
                if (string.IsNullOrEmpty(text)) return;
                Size avail = new Size(cell.Size.Width - cell.InheritedStyle.Padding.Horizontal - 2, cell.Size.Height);
                Size need = TextRenderer.MeasureText(text, cell.InheritedStyle.Font ?? _grid.Font);
                if (need.Width <= avail.Width) return;
                e.ToolTipText = WrapForTip(text, 110);
            };

            MakeCol("#", 40);
            MakeCol("单元格", 110);
            MakeCol("工作表", 140, 220);
            MakeCol("工作簿", 150, 260);
            MakeCol("错误值", 84);
            MakeCol("原因", 150);          // 「公式含错误引用」7 汉字完整显示
            MakeCol("深度", 56);
            MakeCol("公式", 300, 400);
            MakeCol("传播路径", 380);      // 最后一列 Fill 补满窗口
            _grid.Columns[_grid.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Ui.StyleGrid(_grid);
            _grid.Columns[7].DefaultCellStyle.Font = Ui.MonoFont(10f, Font);

            Panel body = new Panel(); body.BackColor = Color.White;
            body.Controls.Add(_grid);
            _root.Controls.Add(Ui.Card(body, "错误来源列表"), 0, 1);

            _lblTip = Ui.Label("", 9f, false); _lblTip.ForeColor = Ui.Green;
            _root.Controls.Add(_lblTip, 0, 2);

            FlowLayoutPanel bar = Ui.ActionBar();
            _btnJump = Ui.Button("跳转", 96); Ui.StylePrimary(_btnJump, false);
            _btnJump.Click += delegate { JumpToSelected(); }; bar.Controls.Add(_btnJump);
            _btnCopyFormula = Ui.Button("复制公式", 110);
            _btnCopyFormula.Click += delegate { CopySelected(false); }; bar.Controls.Add(_btnCopyFormula);
            _btnCopyPath = Ui.Button("复制传播路径", 130);
            _btnCopyPath.Click += delegate { CopySelected(true); }; bar.Controls.Add(_btnCopyPath);
            _btnClose = Ui.Button("关闭", 96); _btnClose.Click += delegate { Close(); }; bar.Controls.Add(_btnClose);
            _root.Controls.Add(bar, 0, 3); Controls.Add(_root);
        }

        private void MakeCol(string header, int width, int maxWidth = 0)
        {
            DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
            c.HeaderText = header;
            c.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            c.MinimumWidth = 50;
            c.SortMode = DataGridViewColumnSortMode.NotSortable;
            Font hf = _grid.ColumnHeadersDefaultCellStyle.Font ?? _grid.Font;
            int need = TextRenderer.MeasureText(header, hf).Width + 32;
            c.Width = Math.Max(width, need);
            if (maxWidth > 0) c.Tag = Math.Max(maxWidth, need);
            _grid.Columns.Add(c);
        }

        // ---------------- 数据填充 ----------------

        private void FillRows()
        {
            _grid.Rows.Clear();
            for (int i = 0; i < _result.Sources.Count; i++)
            {
                ErrorSourceItem item = _result.Sources[i];
                int r = _grid.Rows.Add(
                    i + 1,
                    item.CellAddressA1,
                    item.SheetName,
                    item.WorkbookName,
                    item.ErrorText,
                    ReasonText(item.Reason),
                    item.Depth,
                    OneLine(item.Formula),
                    item.PathText);
                _grid.Rows[r].Tag = item;
                // 错误值列红字
                _grid.Rows[r].Cells[4].Style.ForeColor = Ui.Danger;
            }
            if (_grid.Rows.Count > 0) _grid.Rows[0].Selected = true;
            Ui.AutoFitColumns(_grid);
        }

        private static string ReasonText(ErrorSourceReason reason)
        {
            switch (reason)
            {
                case ErrorSourceReason.LiteralInFormula: return "公式含错误引用";
                case ErrorSourceReason.SelfComputed: return "自身计算错误";
                default: return "";
            }
        }

        // ---------------- 跳转 ----------------

        private void JumpToSelected()
        {
            if (DialogPreviewFactory.IsPreview(this)) return;
            if (_grid.SelectedRows.Count == 0) return;
            ErrorSourceItem item = _grid.SelectedRows[0].Tag as ErrorSourceItem;
            if (item == null) return;

            object flash; int prev;
            string err = PrecedentTrackerEngine.JumpToCellEx(_app,
                item.WorkbookName, item.WorkbookPath, item.SheetName, item.CellAddressA1,
                item.IsCrossWorkbook, out flash, out prev);
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

        private void ClearPendingFlash()
        {
            if (_flashTimer != null) _flashTimer.Stop();
            if (_flashRange != null)
            {
                PrecedentTrackerEngine.ClearFlash(_flashRange, _flashPrev);
                _flashRange = null;
            }
        }

        // ---------------- 复制 ----------------

        private void CopySelected(bool path)
        {
            if (_grid.SelectedRows.Count == 0) { ShowTip("请先在列表里选中一行。"); return; }
            ErrorSourceItem item = _grid.SelectedRows[0].Tag as ErrorSourceItem;
            if (item == null) return;
            string s = path ? item.PathText : item.Formula;
            if (string.IsNullOrEmpty(s)) { ShowTip(path ? "该行没有传播路径。" : "该行没有公式。"); return; }
            for (int i = 0; i < 3; i++)
            {
                try { Clipboard.SetDataObject(s, true); ShowTip((path ? "已复制传播路径：" : "已复制公式：") + TrimForTip(s)); return; }
                catch { System.Threading.Thread.Sleep(80); }
            }
            ShowTip("剪贴板写入失败，请重试。");
        }

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

        private static string TrimForTip(string s)
        {
            const int Max = 46;
            if (string.IsNullOrEmpty(s) || s.Length <= Max) return s ?? "";
            return s.Substring(0, 28) + "…" + s.Substring(s.Length - 15);
        }

        private static string OneLine(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\r", " ").Replace("\n", " ");
        }

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
    }
}
