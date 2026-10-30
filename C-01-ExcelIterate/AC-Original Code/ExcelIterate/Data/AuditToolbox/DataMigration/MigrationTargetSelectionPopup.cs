using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // The "选择" header carries an Excel-style filter button: a small framed
    // square with a down arrow, which turns into a blue funnel while some
    // targets are excluded, just like an active AutoFilter.
    internal sealed class MigrationSelectionHeader : DataGridViewColumnHeaderCell
    {
        internal bool HasExcludedTargets;
        internal static Rectangle FilterButtonBounds(DataGridView grid, Rectangle cellBounds)
        {
            int size = Math.Max(Ui.Px(grid, 12), Math.Min(Ui.Px(grid, 18), cellBounds.Height - Ui.Px(grid, 8)));
            return new Rectangle(cellBounds.Right - Ui.Px(grid, 5) - size, cellBounds.Top + (cellBounds.Height - size) / 2, size, size);
        }
        protected override void Paint(Graphics graphics, Rectangle clipBounds, Rectangle cellBounds, int rowIndex,
            DataGridViewElementStates state, object value, object formattedValue, string errorText,
            DataGridViewCellStyle style, DataGridViewAdvancedBorderStyle borders, DataGridViewPaintParts parts)
        {
            base.Paint(graphics, clipBounds, cellBounds, rowIndex, state, value, formattedValue, errorText, style, borders,
                parts & ~DataGridViewPaintParts.ContentForeground);
            if ((parts & DataGridViewPaintParts.ContentForeground) == 0) return;
            var button = FilterButtonBounds(DataGridView, cellBounds);
            int gap = Ui.Px(DataGridView, 4);
            var label = new Rectangle(cellBounds.Left + gap, cellBounds.Top, Math.Max(0, button.Left - gap - cellBounds.Left - gap), cellBounds.Height);
            TextRenderer.DrawText(graphics, Convert.ToString(formattedValue), style.Font ?? DataGridView.Font, label, style.ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
            bool enabled = DataGridView.Enabled;
            Color ink = !enabled ? SystemColors.GrayText : HasExcludedTargets ? Ui.Blue : style.ForeColor;
            var smoothing = graphics.SmoothingMode;
            try
            {
                using (var fill = new SolidBrush(HasExcludedTargets && enabled ? Color.FromArgb(232, 240, 254) : Color.White))
                    graphics.FillRectangle(fill, button);
                using (var pen = new Pen(HasExcludedTargets && enabled ? Ui.Blue : Color.FromArgb(190, 199, 214)))
                    graphics.DrawRectangle(pen, button.X, button.Y, button.Width - 1, button.Height - 1);
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                float cx = button.Left + (button.Width - 1) / 2f, cy = button.Top + (button.Height - 1) / 2f, u = button.Width / 16f;
                using (var brush = new SolidBrush(ink))
                {
                    if (HasExcludedTargets)
                        // Funnel: a wide top narrowing into a short stem.
                        graphics.FillPolygon(brush, new[] {
                            new PointF(cx - 5f * u, cy - 4.5f * u), new PointF(cx + 5f * u, cy - 4.5f * u), new PointF(cx + 1.2f * u, cy + 0.2f * u),
                            new PointF(cx + 1.2f * u, cy + 4.8f * u), new PointF(cx - 1.2f * u, cy + 3.4f * u), new PointF(cx - 1.2f * u, cy + 0.2f * u) });
                    else
                        graphics.FillPolygon(brush, new[] { new PointF(cx - 4f * u, cy - 2f * u), new PointF(cx + 4f * u, cy - 2f * u), new PointF(cx, cy + 2.5f * u) });
                }
            }
            finally { graphics.SmoothingMode = smoothing; }
        }
    }

    // A lightweight dropdown laid out like Excel's AutoFilter menu: search box,
    // a framed list headed by "(全选)", then 确定/取消 and a resize grip. It is
    // not a second modeless Form. All edits remain a draft until confirmation;
    // searching never changes the main grid's rows.
    internal sealed class MigrationTargetSelectionPopup : ToolStripDropDown
    {
        private sealed class Choice
        {
            internal DataGridViewRow Row;
            internal string Path, Sheet;
            internal bool Selected;
            public override string ToString() { return System.IO.Path.GetFileName(Path) + (Sheet.Length == 0 ? "" : "  ·  " + Sheet); }
        }

        // 1px framed white panel; the frame turns blue while a child has focus.
        private sealed class FramePanel : Panel
        {
            private Color border;
            internal Color Border { get { return border; } set { if (border != value) { border = value; Invalidate(); } } }
            internal FramePanel(Color color)
            {
                border = color; BackColor = Color.White; Padding = new Padding(1); Margin = Padding.Empty;
                SetStyle(ControlStyles.ResizeRedraw | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                using (var pen = new Pen(border)) e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
            }
        }

        // Magnifier while the search is empty; a clickable × once text is typed.
        private sealed class SearchGlyph : Control
        {
            private bool clear;
            internal bool ClearMode { get { return clear; } set { if (clear != value) { clear = value; Cursor = value ? Cursors.Hand : Cursors.Default; Invalidate(); } } }
            internal SearchGlyph()
            {
                BackColor = Color.White; TabStop = false;
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics; g.Clear(BackColor); g.SmoothingMode = SmoothingMode.AntiAlias;
                float s = Math.Min(Width, Height), u = s / 16f, cx = Width / 2f, cy = Height / 2f;
                using (var pen = new Pen(clear ? Ui.Ink : Ui.InkDim, Math.Max(1f, 1.4f * u)))
                {
                    if (clear)
                    {
                        g.DrawLine(pen, cx - 3.5f * u, cy - 3.5f * u, cx + 3.5f * u, cy + 3.5f * u);
                        g.DrawLine(pen, cx + 3.5f * u, cy - 3.5f * u, cx - 3.5f * u, cy + 3.5f * u);
                    }
                    else
                    {
                        float r = 4f * u, ox = cx - 1.2f * u, oy = cy - 1.2f * u;
                        g.DrawEllipse(pen, ox - r, oy - r, r * 2, r * 2);
                        g.DrawLine(pen, ox + r * 0.72f, oy + r * 0.72f, cx + 5.5f * u, cy + 5.5f * u);
                    }
                }
            }
        }

        // Bottom-right drag handle, like the one on Excel's filter menu.
        private sealed class SizeGrip : Control
        {
            private readonly MigrationTargetSelectionPopup owner;
            private Point start; private Size origin; private bool dragging;
            internal SizeGrip(MigrationTargetSelectionPopup popup)
            {
                owner = popup; Cursor = Cursors.SizeNWSE; BackColor = Color.White; TabStop = false;
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.Clear(BackColor);
                int cell = Math.Max(2, Width / 3);
                using (var brush = new SolidBrush(Color.FromArgb(165, 174, 190)))
                    for (int r = 0; r < 3; r++)
                        for (int c = 0; r + c < 3; c++)
                            e.Graphics.FillRectangle(brush, Width - (c + 1) * cell, Height - (r + 1) * cell, cell - 1, cell - 1);
            }
            protected override void OnMouseDown(MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left) { dragging = true; start = Cursor.Position; origin = owner.Size; Capture = true; }
                base.OnMouseDown(e);
            }
            protected override void OnMouseMove(MouseEventArgs e)
            {
                if (dragging && Capture)
                {
                    var p = Cursor.Position;
                    owner.ResizeTo(new Size(origin.Width + p.X - start.X, origin.Height + p.Y - start.Y));
                }
                base.OnMouseMove(e);
            }
            protected override void OnMouseUp(MouseEventArgs e) { dragging = false; Capture = false; base.OnMouseUp(e); }
            protected override void OnMouseCaptureChanged(EventArgs e) { dragging = false; base.OnMouseCaptureChanged(e); }
        }

        private static class Native
        {
            internal const int EM_SETCUEBANNER = 0x1501;
            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
        }

        private static readonly Color OuterFrame = Color.FromArgb(160, 170, 186);
        private static readonly Color InnerFrame = Color.FromArgb(205, 212, 224);
        // Last size the user dragged to, in 96-DPI units; shared within the session.
        private static Size rememberedSize96 = Size.Empty;

        private readonly List<Choice> choices = new List<Choice>();
        private readonly List<Choice> matching = new List<Choice>();
        private readonly TextBox search = new TextBox { Name = "migrationSelectionSearch", BorderStyle = BorderStyle.None };
        private readonly CheckedListBox list = new CheckedListBox { Name = "migrationSelectionList", Dock = DockStyle.Fill, CheckOnClick = true, IntegralHeight = false, BorderStyle = BorderStyle.None, HorizontalScrollbar = true };
        private readonly CheckBox all = new CheckBox { Name = "migrationSelectionAll", Text = "(全选)", AutoSize = true, AutoCheck = false, BackColor = Color.White };
        private readonly LinkLabel invert = new LinkLabel { Name = "migrationSelectionInvert", Text = "反选", AutoSize = true, LinkBehavior = LinkBehavior.HoverUnderline, BackColor = Color.White };
        private readonly Label count = new Label { Name = "migrationSelectionCount", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Ui.InkDim };
        private readonly Label empty = new Label { Name = "migrationSelectionEmpty", Text = "没有匹配的工作簿", AutoSize = false, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Ui.InkDim, BackColor = Color.White, Visible = false };
        private readonly SearchGlyph glyph = new SearchGlyph { Name = "migrationSelectionGlyph" };
        private readonly ToolTip tips = new ToolTip();
        private readonly Action<IDictionary<DataGridViewRow, bool>> apply;
        private readonly FramePanel root;
        private readonly ToolStripControlHost host;
        private readonly Size minimum;
        private readonly float scale;
        private bool populating;
        internal TextBox Search { get { return search; } }
        internal CheckedListBox List { get { return list; } }
        internal int MatchCount { get { return matching.Count; } }

        internal MigrationTargetSelectionPopup(DataGridView grid, Action<IDictionary<DataGridViewRow, bool>> onApply)
        {
            Name = "migrationSelectionPopup"; apply = onApply;
            AutoSize = false; Padding = Padding.Empty; Margin = Padding.Empty; BackColor = Color.White; DropShadowEnabled = true;
            Font = grid.Font;
            using (var g = grid.CreateGraphics()) scale = g.DpiX / 96f;
            foreach (DataGridViewRow row in grid.Rows)
                if (!row.IsNewRow) choices.Add(new Choice { Row = row, Path = Convert.ToString(row.Cells["path"].Value), Sheet = Convert.ToString(row.Cells["sheet"].Value), Selected = DataMigrationForm.IsIncluded(row) });

            foreach (Control control in new Control[] { search, list, all, invert, count, empty }) control.Font = grid.Font;
            minimum = new Size(P(260), P(250));
            int columnWidth = grid.Columns.Contains("include") ? grid.Columns["include"].Width : 0;
            Size size = rememberedSize96.IsEmpty ? new Size(Math.Max(P(320), columnWidth), P(360)) : new Size(P(rememberedSize96.Width), P(rememberedSize96.Height));
            var area = Screen.FromControl(grid).WorkingArea;
            size = new Size(Math.Max(minimum.Width, Math.Min(size.Width, area.Width - P(24))), Math.Max(minimum.Height, Math.Min(size.Height, area.Height - P(24))));

            root = new FramePanel(OuterFrame) { Name = "migrationSelectionFrame", Size = size, Font = grid.Font };
            var layout = new TableLayoutPanel { Name = "migrationSelectionPanel", Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = Color.White, Padding = new Padding(P(8), P(8), P(8), P(8)), Margin = Padding.Empty };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            int searchHeight = Math.Max(P(28), search.PreferredHeight + P(10));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, searchHeight + P(6)));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, P(30) + P(8)));

            // Search box: borderless TextBox + glyph inside a framed strip.
            var searchFrame = new FramePanel(InnerFrame) { Name = "migrationSelectionSearchFrame", Dock = DockStyle.Top, Height = searchHeight, Margin = new Padding(0, 0, 0, P(6)) };
            search.BackColor = Color.White;
            searchFrame.Controls.Add(search); searchFrame.Controls.Add(glyph);
            searchFrame.Layout += delegate
            {
                int inner = searchFrame.Height - 2, glyphWidth = P(24);
                glyph.SetBounds(searchFrame.Width - 1 - glyphWidth, 1, glyphWidth, inner);
                search.SetBounds(1 + P(7), 1 + (inner - search.Height) / 2, Math.Max(P(20), searchFrame.Width - 2 - P(7) - glyphWidth), search.Height);
            };
            layout.Controls.Add(searchFrame, 0, 0);

            // List: "(全选)" strip, a hairline, then the workbook choices.
            var listFrame = new FramePanel(InnerFrame) { Name = "migrationSelectionListFrame", Dock = DockStyle.Fill, Margin = Padding.Empty };
            var strip = new Panel { Name = "migrationSelectionStrip", Dock = DockStyle.Top, Height = Math.Max(P(28), all.PreferredSize.Height + P(8)), BackColor = Color.White };
            strip.Controls.Add(all); strip.Controls.Add(invert);
            strip.Layout += delegate
            {
                var allSize = all.PreferredSize; var invertSize = invert.PreferredSize;
                invert.SetBounds(strip.Width - P(10) - invertSize.Width, (strip.Height - invertSize.Height) / 2, invertSize.Width, invertSize.Height);
                all.SetBounds(P(5), (strip.Height - allSize.Height) / 2, Math.Min(allSize.Width, Math.Max(0, invert.Left - P(12) - P(5))), allSize.Height);
            };
            var rule = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Ui.GridLine };
            var listHost = new Panel { Name = "migrationSelectionListHost", Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(P(4), P(3), 0, 0) };
            list.BackColor = Color.White; list.ForeColor = Ui.Ink;
            listHost.Controls.Add(list); listHost.Controls.Add(empty); empty.BringToFront();
            listHost.Layout += delegate { empty.SetBounds(listHost.Padding.Left, listHost.Padding.Top, Math.Max(0, listHost.ClientSize.Width - listHost.Padding.Horizontal), Math.Max(0, Math.Min(P(48), listHost.ClientSize.Height - listHost.Padding.Vertical))); };
            // Docking runs from the highest index down: strip, rule, then Fill.
            listFrame.Controls.Add(listHost); listFrame.Controls.Add(rule); listFrame.Controls.Add(strip);
            listHost.BringToFront();
            layout.Controls.Add(listFrame, 0, 1);

            // Count on the left, 确定 (default) and 取消 on the right.
            var actions = new TableLayoutPanel { Name = "migrationSelectionActions", Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, Margin = new Padding(0, P(8), 0, 0), Padding = Padding.Empty, BackColor = Color.White };
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); actions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); actions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            actions.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var confirm = SmallButton("确定", true); confirm.Name = "migrationSelectionApply";
            var cancel = SmallButton("取消", false); cancel.Name = "migrationSelectionCancel";
            count.Dock = DockStyle.Fill; count.Margin = Padding.Empty;
            actions.Controls.Add(count, 0, 0); actions.Controls.Add(confirm, 1, 0); actions.Controls.Add(cancel, 2, 0);
            layout.Controls.Add(actions, 0, 2);

            root.Controls.Add(layout);
            var grip = new SizeGrip(this) { Name = "migrationSelectionGrip", Size = new Size(P(10), P(10)) };
            root.Controls.Add(grip); grip.BringToFront();
            root.Layout += delegate { grip.Location = new Point(root.Width - 1 - grip.Width, root.Height - 1 - grip.Height); };

            host = new ToolStripControlHost(root) { AutoSize = false, Size = root.Size, Margin = Padding.Empty, Padding = Padding.Empty };
            Items.Add(host); Size = root.Size;

            search.TextChanged += delegate { glyph.ClearMode = search.TextLength > 0; Populate(); };
            search.GotFocus += delegate { searchFrame.Border = Ui.Blue; };
            search.LostFocus += delegate { searchFrame.Border = InnerFrame; };
            search.HandleCreated += delegate { Native.SendMessage(search.Handle, Native.EM_SETCUEBANNER, (IntPtr)1, "搜索文件名、路径或工作表"); };
            glyph.Click += delegate { if (search.TextLength > 0) { search.Clear(); } search.Focus(); };
            searchFrame.Click += delegate { search.Focus(); };
            all.Click += delegate { SelectMatching(all.CheckState != CheckState.Checked); };
            invert.LinkClicked += delegate { InvertMatching(); };
            list.ItemCheck += delegate(object sender, ItemCheckEventArgs e)
            { if (!populating) { matching[e.Index].Selected = e.NewValue == CheckState.Checked; UpdateCount(); } };
            list.MouseMove += delegate(object sender, MouseEventArgs e)
            { int index = list.IndexFromPoint(e.Location); string text = index < 0 || index >= matching.Count ? "" : matching[index].Path; if (tips.GetToolTip(list) != text) tips.SetToolTip(list, text); };
            list.KeyDown += delegate(object sender, KeyEventArgs e)
            { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; ApplySelection(); } };
            confirm.Click += delegate { ApplySelection(); }; cancel.Click += delegate { Close(); };
            Opened += delegate { BeginInvoke(new Action(delegate { if (!IsDisposed && Visible) search.Focus(); })); };
            search.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; ApplySelection(); }
                else if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; if (search.TextLength > 0) search.Clear(); else Close(); }
                else if (e.KeyCode == Keys.Down && list.Items.Count > 0) { e.SuppressKeyPress = true; list.Focus(); if (list.SelectedIndex < 0) list.SelectedIndex = 0; }
            };
            Populate();
        }
        private int P(int value) { return (int)Math.Round(value * scale); }
        private RoundedButton SmallButton(string text, bool primary)
        {
            var button = Ui.Button(text, 72);
            if (primary) Ui.StylePrimary(button);
            button.Font = Font; button.BackColor = Color.White; button.Radius = 6;
            button.MinimumSize = Size.Empty; button.Size = new Size(P(72), P(30));
            button.Margin = new Padding(P(8), 0, 0, 0); button.Anchor = AnchorStyles.Right;
            return button;
        }
        internal void ResizeTo(Size size)
        {
            var area = Screen.FromPoint(Bounds.Location).WorkingArea;
            size = new Size(Math.Max(minimum.Width, Math.Min(size.Width, area.Right - Left)), Math.Max(minimum.Height, Math.Min(size.Height, area.Bottom - Top)));
            if (size == Size) return;
            SuspendLayout();
            try { root.Size = size; host.Size = size; Size = size; }
            finally { ResumeLayout(true); }
            rememberedSize96 = new Size((int)Math.Round(size.Width / scale), (int)Math.Round(size.Height / scale));
        }
        private void Populate()
        {
            string term = search.Text.Trim(); populating = true; list.BeginUpdate();
            try
            {
                matching.Clear(); list.Items.Clear();
                foreach (var choice in choices)
                    if (term.Length == 0 || choice.Path.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 || choice.Sheet.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0)
                    { matching.Add(choice); list.Items.Add(choice, choice.Selected); }
            }
            finally { list.EndUpdate(); populating = false; }
            all.Text = term.Length == 0 ? "(全选)" : "(选择所有搜索结果)";
            empty.Visible = matching.Count == 0;
            UpdateCount();
        }
        internal void SelectMatching(bool selected)
        { foreach (var choice in matching) choice.Selected = selected; Populate(); }
        internal void InvertMatching()
        { foreach (var choice in matching) choice.Selected = !choice.Selected; Populate(); }
        private void UpdateCount()
        {
            if (IsDisposed) return;
            int selected = matching.Count(c => c.Selected);
            all.Enabled = invert.Enabled = matching.Count > 0;
            all.CheckState = matching.Count == 0 || selected == 0 ? CheckState.Unchecked : selected == matching.Count ? CheckState.Checked : CheckState.Indeterminate;
            string text = "已选 " + choices.Count(c => c.Selected) + " / " + choices.Count;
            if (matching.Count != choices.Count) text += "，显示 " + matching.Count;
            count.Text = text;
        }
        internal void ApplySelection()
        {
            apply(choices.ToDictionary(c => c.Row, c => c.Selected)); Close();
        }
        protected override void Dispose(bool disposing)
        { if (disposing) tips.Dispose(); base.Dispose(disposing); }
    }
}
