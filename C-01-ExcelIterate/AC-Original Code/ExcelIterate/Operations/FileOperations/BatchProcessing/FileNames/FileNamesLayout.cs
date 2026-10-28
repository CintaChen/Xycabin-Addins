using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed partial class FileNamesForm
    {
        private readonly HashSet<string> manualColumnWidths = new HashSet<string>(StringComparer.Ordinal);
        private bool fittingColumns, columnsReady, measuringOptions;
        private TableLayoutPanel fileLayoutRoot;
        private const TextFormatFlags ColumnTextFlags = TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine;

        internal void BeginColumnFitCycle() { manualColumnWidths.Clear(); }
        internal void FitResultColumns()
        {
            if (fittingColumns || grid.IsDisposed || grid.Columns.Count == 0) return;
            fittingColumns = true;
            grid.SuspendLayout();
            try
            {
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                grid.AllowUserToResizeColumns = true;
                grid.ScrollBars = ScrollBars.Both;
                Ui.EnsureGridHeaders(grid);
                foreach (DataGridViewColumn column in grid.Columns)
                {
                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    column.Resizable = DataGridViewTriState.True;
                    if (manualColumnWidths.Contains(column.Name)) continue;
                    Font body = column.InheritedStyle.Font ?? grid.Font;
                    int padding = column.InheritedStyle.Padding.Horizontal;
                    int cap = Math.Max(column.MinimumWidth, TextRenderer.MeasureText(new string('0', 30), body, Size.Empty, ColumnTextFlags).Width + padding);
                    int need = column.MinimumWidth;
                    var button = column as DataGridViewButtonColumn;
                    if (button != null)
                        need = Math.Max(need, TextRenderer.MeasureText(button.Text, body, Size.Empty, ColumnTextFlags).Width + padding + Ui.Px(grid, 12));
                    else
                    {
                        foreach (DataGridViewRow row in grid.Rows)
                        {
                            if (row.IsNewRow) continue;
                            int width = TextRenderer.MeasureText(Convert.ToString(row.Cells[column.Index].FormattedValue), row.Cells[column.Index].InheritedStyle.Font ?? body, Size.Empty, ColumnTextFlags).Width + padding;
                            need = Math.Max(need, Math.Min(width, cap));
                            if (need >= cap) break;
                        }
                        need = Math.Min(need, cap);
                    }
                    column.Width = Math.Max(column.MinimumWidth, need);
                }
                columnsReady = true;
            }
            finally { grid.ResumeLayout(); fittingColumns = false; }
        }

        private void ConfigureFileNamesLayout(TableLayoutPanel root)
        {
            fileLayoutRoot = root;
            foreach (Control control in DialogFormatting.Walk(options))
            {
                var button = control as Button;
                if (button != null) button.Margin = new Padding(button.Margin.Left, button.Margin.Top, button.Margin.Right, Math.Min(button.Margin.Bottom, 4));
            }
            options.GetControlFromPosition(0, 1).Margin = new Padding(3, 4, 6, 4);
            explanation.Margin = new Padding(3, 4, 6, 4);
            grid.ColumnWidthChanged += delegate(object sender, DataGridViewColumnEventArgs e)
            {
                if (columnsReady && !fittingColumns && grid.IsHandleCreated) manualColumnWidths.Add(e.Column.Name);
            };
            root.Layout += delegate { ReflowFileOptions(); };
            root.SizeChanged += delegate { ReflowFileOptions(); };
            explanation.TextChanged += delegate { ReflowFileOptions(); };
            explanation.FontChanged += delegate { ReflowFileOptions(); };
            Load += delegate { ReflowFileOptions(); FitResultColumns(); };
            ReflowFileOptions();
        }
        private void ReflowFileOptions()
        {
            if (measuringOptions || fileLayoutRoot == null || fileLayoutRoot.IsDisposed) return;
            measuringOptions = true;
            try
            {
                int width = Math.Max(1, fileLayoutRoot.ClientSize.Width - fileLayoutRoot.Padding.Horizontal - options.Margin.Horizontal);
                options.MaximumSize = new Size(width, 0);
                options.Width = width;
                foreach (Control row in options.Controls)
                {
                    var flow = row as FlowLayoutPanel;
                    if (flow == null) continue;
                    int available = Math.Max(1, width - options.Padding.Horizontal - flow.Margin.Horizontal);
                    flow.MaximumSize = new Size(available, 0);
                    flow.Width = available;
                    flow.PerformLayout();
                }
                int labelWidth = Math.Max(1, width - options.Padding.Horizontal - explanation.Margin.Horizontal);
                if (!DialogControlSizing.HasWidth(explanation)) explanation.MaximumSize = new Size(labelWidth, 0);
                int total = options.Padding.Vertical;
                foreach (Control row in options.Controls)
                {
                    int index = options.GetRow(row);
                    int available = Math.Max(1, width - options.Padding.Horizontal - row.Margin.Horizontal);
                    int height = row.AutoSize ? row.GetPreferredSize(new Size(available, 0)).Height : row.Height;
                    if (row == explanation && !DialogControlSizing.HasHeight(explanation))
                    {
                        int wrap = DialogControlSizing.HasWidth(explanation) ? explanation.Width : labelWidth;
                        height = TextRenderer.MeasureText(explanation.Text, explanation.Font, new Size(Math.Max(1, wrap), int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + explanation.Padding.Vertical;
                        if (!explanation.AutoSize) explanation.Height = height;
                    }
                    height = Math.Max(row.MinimumSize.Height, Math.Max(1, height)) + row.Margin.Vertical;
                    // TableLayoutPanel can retain extra height in its last AutoSize row.
                    // Recompute every row from its current width rather than reusing that allocation.
                    options.RowStyles[index].SizeType = SizeType.Absolute;
                    options.RowStyles[index].Height = height;
                    total += height;
                }
                options.Height = total;
                options.PerformLayout();
            }
            finally { measuringOptions = false; }
        }
    }
}
