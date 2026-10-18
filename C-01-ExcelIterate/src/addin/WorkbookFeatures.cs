using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class ComBag : IDisposable
    {
        private readonly List<object> references = new List<object>();
        internal object Own(object value) { if (value != null) references.Add(value); return value; }

        /// <summary>
        /// 把引用移出本袋：调用方将长期持有它，本袋 Dispose 时不得 Release。
        /// 用于「方法内取到、但要返回给上层继续用」的 COM 引用（Workbook / Worksheet）。
        /// 少了这一步，返回的引用会在 using 结束时被 Release，之后再碰就是
        /// “COM 对象与其基础 RCW 分开后就不能再使用”。
        /// </summary>
        internal object Detach(object value)
        {
            if (value != null)
                for (int i = references.Count - 1; i >= 0; i--)
                    if (ReferenceEquals(references[i], value)) { references.RemoveAt(i); break; }
            return value;
        }

        public void Dispose()
        {
            List<string> errors = new List<string>();
            for (int i = references.Count - 1; i >= 0; i--)
                try { if (Marshal.IsComObject(references[i])) Marshal.ReleaseComObject(references[i]); } catch (Exception ex) { errors.Add(ex.Message); }
            references.Clear();
            if (errors.Count != 0) throw new InvalidOperationException("引用释放失败：" + string.Join("；", errors));
        }
    }
    internal sealed class HostQuiet : IDisposable
    {
        private readonly object app;
        private readonly bool events, screen;
        private bool disposed;
        internal HostQuiet(object application)
        {
            app = application;
            events = ComFormulaHost.Boolean(((dynamic)app).EnableEvents, "EnableEvents");
            screen = ComFormulaHost.Boolean(((dynamic)app).ScreenUpdating, "ScreenUpdating");
            try { ((dynamic)app).EnableEvents = false; ((dynamic)app).ScreenUpdating = false;
                if (ComFormulaHost.Boolean(((dynamic)app).EnableEvents, "EnableEvents") || ComFormulaHost.Boolean(((dynamic)app).ScreenUpdating, "ScreenUpdating")) throw new InvalidOperationException("宿主未接受暂停事件或刷新。"); }
            catch { Dispose(); throw; }
        }
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            List<string> errors = new List<string>();
            try { ((dynamic)app).ScreenUpdating = screen; if (ComFormulaHost.Boolean(((dynamic)app).ScreenUpdating, "ScreenUpdating") != screen) throw new Exception("读回不一致"); } catch (Exception ex) { errors.Add("刷新恢复：" + ex.Message); }
            try { ((dynamic)app).EnableEvents = events; if (ComFormulaHost.Boolean(((dynamic)app).EnableEvents, "EnableEvents") != events) throw new Exception("读回不一致"); } catch (Exception ex) { errors.Add("事件恢复：" + ex.Message); }
            if (errors.Count != 0) throw new InvalidOperationException(string.Join("；", errors));
        }
    }
    internal static class WorkbookUi
    {
        internal static string Quote(string name) { return "'" + name.Replace("'", "''") + "'"; }
        internal static string UniqueName(string basis, IEnumerable<string> existing)
        {
            HashSet<string> names = new HashSet<string>(existing, StringComparer.OrdinalIgnoreCase);
            if (!names.Contains(basis)) return basis;
            for (int i = 2; ; i++) { string suffix = "(" + i + ")"; string name = basis.Substring(0, Math.Min(basis.Length, 31 - suffix.Length)) + suffix; if (!names.Contains(name)) return name; }
        }
        internal static object NewSheet(object workbook, string basis, ComBag bag, bool first)
        {
            if (ComFormulaHost.Boolean(((dynamic)workbook).ProtectStructure, "ProtectStructure")) throw new InvalidOperationException("工作簿结构受到保护，不能新增工作表。");
            dynamic sheets = bag.Own(((dynamic)workbook).Sheets); List<string> names = new List<string>();
            for (int i = 1; i <= Convert.ToInt32(sheets.Count); i++) { using (ComBag item = new ComBag()) { dynamic s = item.Own(sheets.Item(i)); names.Add((string)s.Name); } }
            dynamic worksheets = bag.Own(((dynamic)workbook).Worksheets);
            object anchor = bag.Own(sheets.Item(first ? 1 : Convert.ToInt32(sheets.Count)));
            dynamic result = bag.Own(first ? worksheets.Add(anchor, Type.Missing, 1, Type.Missing) : worksheets.Add(Type.Missing, anchor, 1, Type.Missing));
            try { result.Name = UniqueName(basis, names); return result; }
            catch { object app = bag.Own(((dynamic)workbook).Application); DeleteSheet(app, (object)result); throw; }
        }
        internal static void DeleteSheet(object app, object sheet)
        {
            bool alerts = ComFormulaHost.Boolean(((dynamic)app).DisplayAlerts, "DisplayAlerts");
            try { ((dynamic)app).DisplayAlerts = false; ((dynamic)sheet).Delete(); }
            finally { ((dynamic)app).DisplayAlerts = alerts; }
        }
        internal static void StyleHeader(object sheet, int startColumn, int columns)
        { StyleHeader(sheet, 1, startColumn, columns); }

        /// <summary>表头行样式（row 为表头所在行，抵消分录可从任意起始单元格写起）。</summary>
        internal static void StyleHeader(object sheet, int row, int startColumn, int columns)
        {
            using (ComBag bag = new ComBag())
            {
                string first = ComFormulaHost.Address(row, startColumn);
                string last = ComFormulaHost.Address(row, startColumn + columns - 1);
                dynamic header = bag.Own(((dynamic)sheet).Range[first + ":" + last]);
                dynamic headFont = bag.Own(header.Font), fill = bag.Own(header.Interior);
                headFont.Bold = true; headFont.Color = 0x473224; fill.Color = 0xFBEFF0; header.RowHeight = 26;
                header.HorizontalAlignment = -4108; header.VerticalAlignment = -4108; header.WrapText = false;
            }
        }
        internal static void Center(object range)
        {
            ((dynamic)range).HorizontalAlignment = -4108;
            ((dynamic)range).VerticalAlignment = -4108;
        }
        private static bool PositiveHeight(object value)
        {
            if (value == null) return false;
            double height = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            return !double.IsNaN(height) && !double.IsInfinity(height) && height > 0;
        }
        // Hidden rows can report zero rather than their latent stored height.
        // The caller owns the row reference and has already captured Hidden.
        internal static object ReadRowHeight(object row, bool hidden)
        {
            object height = ((dynamic)row).RowHeight;
            if (PositiveHeight(height)) return height;
            if (!hidden) throw new InvalidOperationException("无法读取有效行高，未开始格式适配。");
            Exception failure = null;
            try
            {
                ((dynamic)row).Hidden = false;
                if (ComFormulaHost.Boolean(((dynamic)row).Hidden, "Hidden")) throw new InvalidOperationException("宿主未接受临时显露输出行。");
                height = ((dynamic)row).RowHeight;
                if (!PositiveHeight(height)) throw new InvalidOperationException("临时显露后仍无法读取有效行高。");
            }
            catch (Exception ex) { failure = ex; }
            finally
            {
                try
                {
                    ((dynamic)row).Hidden = true;
                    if (!ComFormulaHost.Boolean(((dynamic)row).Hidden, "Hidden")) throw new InvalidOperationException("隐藏状态读回不一致");
                }
                catch (Exception ex) { throw new InvalidOperationException("读取隐藏行高后未能恢复原隐藏状态：" + ex.Message + (failure == null ? "" : "；原始失败：" + failure.Message), failure ?? ex); }
            }
            if (failure != null) throw new InvalidOperationException("读取隐藏行高失败：" + failure.Message, failure);
            return height;
        }
        internal static void ClearSheet(object sheet)
        {
            using (ComBag bag = new ComBag())
            {
                try { ((dynamic)sheet).AutoFilterMode = false; } catch {}
                dynamic cells = bag.Own(((dynamic)sheet).Cells); cells.Clear();
            }
        }
        internal static object EnsureSheet(object workbook, string basis, ComBag bag, bool first)
        {
            dynamic worksheets = bag.Own(((dynamic)workbook).Worksheets);
            for (int i = 1; i <= Convert.ToInt32(worksheets.Count); i++)
            {
                dynamic existing = bag.Own(worksheets.Item(i));
                if (string.Equals(Convert.ToString(existing.Name), basis, StringComparison.OrdinalIgnoreCase))
                {
                    ClearSheet((object)existing);
                    return existing;
                }
            }
            return NewSheet(workbook, basis, bag, first);
        }
        internal static void StyleTable(object sheet, int columns, int rows)
        {
            using (ComBag bag = new ComBag())
            {
                string last = ComFormulaHost.Address(1, columns).TrimEnd('1');
                dynamic data = bag.Own(((dynamic)sheet).Range["A1:" + last + Math.Max(1, rows)]);
                dynamic font = bag.Own(data.Font); font.Name = "Microsoft YaHei UI"; font.Size = 10;
            }
            StyleHeader(sheet, 1, columns);
        }
        // Fits only the product-owned output rectangle, after values, number
        // formats, hyperlinks and fonts are final. Hidden rows are temporarily
        // exposed for measurement without changing the filter or its criteria.
        internal static void FitOutput(object sheet, int startColumn, int columns, int rows)
        {
            if (startColumn < 1 || columns < 1 || rows < 1) throw new ArgumentOutOfRangeException("输出区域必须包含表头。");
            using (ComBag bag = new ComBag())
            {
                string address = ComFormulaHost.Address(1, startColumn) + ":" + ComFormulaHost.Address(rows, startColumn + columns - 1);
                dynamic output = bag.Own(((dynamic)sheet).Range[address]);
                // Complete all snapshots before changing visibility or formats.
                var before = new OutputFormatSnapshot(sheet, (object)output, startColumn, columns, rows);
                Exception failure = null;
                var restoration = new List<string>();
                try
                {
                    try
                    {
                    foreach (int r in before.HiddenRows) using (ComBag one = new ComBag())
                    {
                        dynamic cell = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(r, startColumn)]);
                        dynamic row = one.Own(cell.EntireRow); row.Hidden = false;
                        if (ComFormulaHost.Boolean((object)row.Hidden, "Hidden")) throw new InvalidOperationException("宿主未接受临时显露输出行。");
                    }
                    StyleHeader(sheet, startColumn, columns);
                    output.WrapText = false;
                    dynamic widths = bag.Own(output.Columns); widths.AutoFit();
                    bool wrapped = false;
                    for (int c = startColumn; c < startColumn + columns; c++) using (ComBag one = new ComBag())
                    {
                        dynamic column = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(1, c) + ":" + ComFormulaHost.Address(rows, c)]);
                        double width = Convert.ToDouble((object)column.ColumnWidth, CultureInfo.InvariantCulture);
                        if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0) throw new InvalidOperationException("无法读回有效输出列宽。");
                        // Excel-compatible hosts cap ColumnWidth at 255. Keep the
                        // header on one line; wrap only the data at that limit.
                        if (rows > 1 && width >= 255)
                        {
                            dynamic body = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(2, c) + ":" + ComFormulaHost.Address(rows, c)]);
                            body.WrapText = true; wrapped = true;
                        }
                    }
                    if (wrapped)
                    {
                        dynamic body = bag.Own(((dynamic)sheet).Range[ComFormulaHost.Address(2, startColumn) + ":" + ComFormulaHost.Address(rows, startColumn + columns - 1)]);
                        dynamic heights = bag.Own(body.Rows); heights.AutoFit();
                    }
                    dynamic header = bag.Own(((dynamic)sheet).Range[ComFormulaHost.Address(1, startColumn) + ":" + ComFormulaHost.Address(1, startColumn + columns - 1)]);
                    dynamic headerRows = bag.Own(header.Rows); headerRows.AutoFit();
                    object headerHeight = header.RowHeight;
                    if (!PositiveHeight(headerHeight)) throw new InvalidOperationException("无法读回有效表头行高。");
                    if (Convert.ToDouble(headerHeight, CultureInfo.InvariantCulture) < 26) header.RowHeight = 26;
                    }
                    catch (Exception ex) { failure = ex; }
                }
                finally
                {
                    if (failure != null) before.Restore(sheet, (object)output, restoration);
                    int previousErrors = restoration.Count;
                    foreach (int r in before.HiddenRows)
                    {
                        try { using (ComBag one = new ComBag())
                        {
                            dynamic cell = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(r, startColumn)]);
                            dynamic row = one.Own(cell.EntireRow); row.Hidden = true;
                            if (!ComFormulaHost.Boolean((object)row.Hidden, "Hidden")) throw new InvalidOperationException("隐藏状态读回不一致");
                        } }
                        catch (Exception ex) { restoration.Add("第" + r + "行：" + ex.Message); }
                    }
                    // Visibility is part of the same format transaction. A failed
                    // finally restore must undo successful measurement as well.
                    // Retain the initial row diagnostic even if the retry works.
                    if (failure == null && restoration.Count > previousErrors)
                    {
                        failure = new InvalidOperationException("隐藏状态恢复失败，已尝试回滚输出格式。");
                        before.Restore(sheet, (object)output, restoration);
                        foreach (int r in before.HiddenRows)
                        {
                            try { using (ComBag one = new ComBag())
                            {
                                dynamic cell = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(r, startColumn)]);
                                dynamic row = one.Own(cell.EntireRow); row.Hidden = true;
                                if (!ComFormulaHost.Boolean((object)row.Hidden, "Hidden")) throw new InvalidOperationException("隐藏状态读回不一致");
                            } }
                            catch (Exception ex) { restoration.Add("第" + r + "行再次恢复：" + ex.Message); }
                        }
                    }
                }
                if (restoration.Count != 0) throw new InvalidOperationException("输出格式测量后原状态未完整恢复：" + string.Join("；", restoration) + (failure == null ? "" : "；原始失败：" + failure.Message), failure);
                if (failure != null) throw new InvalidOperationException("输出格式适配失败：" + failure.Message, failure);
            }
        }
        private sealed class OutputFormatSnapshot
        {
            private readonly int first, columns, rows;
            private readonly object uniformWrap;
            private readonly bool[,] cellWrap;
            private readonly object[] widths, heights, bold, fontColor, fillColor, horizontal, vertical;
            internal readonly List<int> HiddenRows = new List<int>();
            internal OutputFormatSnapshot(object sheet, object output, int first, int columns, int rows)
            {
                this.first = first; this.columns = columns; this.rows = rows;
                widths = new object[columns]; heights = new object[rows]; bold = new object[columns];
                fontColor = new object[columns]; fillColor = new object[columns]; horizontal = new object[columns]; vertical = new object[columns];
                uniformWrap = ((dynamic)output).WrapText;
                if (!(uniformWrap is bool)) cellWrap = new bool[rows, columns];
                for (int r = 1; r <= rows; r++) using (ComBag one = new ComBag())
                {
                    dynamic cell = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(r, first)]);
                    dynamic row = one.Own(cell.EntireRow);
                    if (ComFormulaHost.Boolean((object)row.Hidden, "Hidden")) HiddenRows.Add(r);
                }
                var hiddenSet = new HashSet<int>(HiddenRows);
                for (int r = 1; r <= rows; r++) using (ComBag one = new ComBag())
                {
                    dynamic cell = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(r, first)]);
                    object row = one.Own(cell.EntireRow); heights[r - 1] = ReadRowHeight(row, hiddenSet.Contains(r));
                }
                for (int c = 0; c < columns; c++) using (ComBag one = new ComBag())
                {
                    dynamic cell = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(1, first + c)]);
                    widths[c] = cell.ColumnWidth;
                    dynamic font = one.Own(cell.Font), fill = one.Own(cell.Interior);
                    bold[c] = font.Bold; fontColor[c] = font.Color; fillColor[c] = fill.Color;
                    horizontal[c] = cell.HorizontalAlignment; vertical[c] = cell.VerticalAlignment;
                }
                if (cellWrap != null) for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++) using (ComBag one = new ComBag())
                {
                    dynamic cell = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(r + 1, first + c)]);
                    cellWrap[r, c] = ComFormulaHost.Boolean((object)cell.WrapText, "WrapText");
                }
            }
            private static void Attempt(Action action, string label, List<string> errors)
            { try { action(); } catch (Exception ex) { errors.Add(label + "：" + ex.Message); } }
            internal void Restore(object sheet, object output, List<string> errors)
            {
                if (cellWrap == null) Attempt(delegate { ((dynamic)output).WrapText = uniformWrap; }, "换行", errors);
                else for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++)
                {
                    int rr = r, cc = c;
                    Attempt(delegate { using (ComBag one = new ComBag())
                    { dynamic cell = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(rr + 1, first + cc)]); cell.WrapText = cellWrap[rr, cc]; } }, "换行" + ComFormulaHost.Address(r + 1, first + c), errors);
                }
                for (int c = 0; c < columns; c++)
                {
                    int cc = c;
                    Attempt(delegate { using (ComBag one = new ComBag())
                    {
                        dynamic cell = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(1, first + cc)]);
                        cell.ColumnWidth = widths[cc];
                    } }, "列宽" + (first + c), errors);
                    Attempt(delegate { using (ComBag one = new ComBag())
                    {
                        dynamic cell = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(1, first + cc)]);
                        Attempt(delegate { cell.HorizontalAlignment = horizontal[cc]; }, "表头水平对齐" + (first + cc), errors);
                        Attempt(delegate { cell.VerticalAlignment = vertical[cc]; }, "表头垂直对齐" + (first + cc), errors);
                        Attempt(delegate { dynamic font = one.Own(cell.Font);
                            Attempt(delegate { font.Bold = bold[cc]; }, "表头粗体" + (first + cc), errors);
                            Attempt(delegate { font.Color = fontColor[cc]; }, "表头字色" + (first + cc), errors);
                        }, "表头字体引用" + (first + cc), errors);
                        Attempt(delegate { dynamic fill = one.Own(cell.Interior);fill.Color = fillColor[cc]; }, "表头底色" + (first + cc), errors);
                    } }, "列及表头" + (first + c), errors);
                }
                for (int r = 0; r < rows; r++)
                {
                    int rr = r;
                    Attempt(delegate { using (ComBag one = new ComBag())
                    { dynamic cell = one.Own(((dynamic)sheet).Range[ComFormulaHost.Address(rr + 1, first)]); dynamic row = one.Own(cell.EntireRow); row.RowHeight = heights[rr]; } }, "行高" + (r + 1), errors);
                }
            }
        }
    }
    internal sealed class FilterCriterion
    {
        internal int Row, Column;
        internal object Value;
    }
    internal static class QuickFilterCriteria
    {
        // Physical row identity defines AND, distinct rows define OR. A union
        // across columns must not create a Cartesian product of selected values.
        internal static object[,] Build(string[] headers, IList<FilterCriterion> cells)
        {
            SortedDictionary<int, Dictionary<int, object>> rows = new SortedDictionary<int, Dictionary<int, object>>();
            SortedSet<int> columns = new SortedSet<int>();
            foreach (FilterCriterion cell in cells)
            {
                if (cell.Value is ErrorWrapper || (cell.Value is int && (((uint)(int)cell.Value & 0xFFFF0000u) == 0x800A0000u)) || (cell.Value is uint && (((uint)cell.Value & 0xFFFF0000u) == 0x800A0000u))) throw new InvalidOperationException("条件不能包含错误值。");
                if (cell.Column < 0 || cell.Column >= headers.Length) throw new InvalidOperationException("条件列不在数据区域内。");
                Dictionary<int, object> row;
                if (!rows.TryGetValue(cell.Row, out row)) rows.Add(cell.Row, row = new Dictionary<int, object>());
                row[cell.Column] = cell.Value; columns.Add(cell.Column);
            }
            List<Dictionary<int, object>> effective = new List<Dictionary<int, object>>();
            foreach (Dictionary<int, object> row in rows.Values)
            {
                bool any = false; foreach (object value in row.Values) if (value != null && Convert.ToString(value, CultureInfo.InvariantCulture).Length != 0) any = true;
                if (any) effective.Add(row); // An entirely blank OR row would match everything.
            }
            if (effective.Count == 0) throw new InvalidOperationException("所选条件均为空，请选择有内容的数据单元格。");
            object[,] result = new object[effective.Count + 1, columns.Count]; int c = 0;
            foreach (int col in columns)
            {
                result[0, c] = headers[col];
                for (int r = 0; r < effective.Count; r++) { object value; if (effective[r].TryGetValue(col, out value)) result[r + 1, c] = value; }
                c++;
            }
            return result;
        }
    }
    internal static class QuickAdvancedFilter
    {
        internal static int Run(object app)
        {
            using (ComBag bag = new ComBag())
            {
                dynamic window = bag.Own(((dynamic)app).ActiveWindow), groups = bag.Own(window.SelectedSheets);
                if (Convert.ToInt32(groups.Count) != 1) throw new InvalidOperationException("请先取消工作表组合选择。");
                dynamic sheet = bag.Own(((dynamic)app).ActiveSheet), selection = bag.Own(((dynamic)app).Selection);
                if (sheet == null || selection == null) throw new InvalidOperationException("请选中数据行中的条件单元格。");
                using (WorksheetIdentity identity = WorksheetIdentity.Capture((object)sheet))
                {
                    object selectedSheet = bag.Own(selection.Worksheet);
                    if (!identity.Matches(selectedSheet)) throw new InvalidOperationException("请在当前工作表中选择条件。");
                    if (ComFormulaHost.Boolean((object)sheet.ProtectContents, "ProtectContents")) throw new InvalidOperationException("请先取消工作表保护，再取得连续数据区域。");
                    dynamic areas = bag.Own(selection.Areas), first = bag.Own(areas.Item(1));
                    dynamic firstCell = bag.Own(sheet.Range[ComFormulaHost.Address(Convert.ToInt32(first.Row), Convert.ToInt32(first.Column))]);
                    dynamic source = bag.Own(firstCell.CurrentRegion);
                    int top = Convert.ToInt32(source.Row), left = Convert.ToInt32(source.Column);
                    dynamic sourceRows = bag.Own(source.Rows), sourceColumns = bag.Own(source.Columns);
                    int height = Convert.ToInt32(sourceRows.Count), width = Convert.ToInt32(sourceColumns.Count);
                    if (height < 2) throw new InvalidOperationException("条件应位于带表头的连续数据表内。");
                    string[] headers = new string[width]; HashSet<string> unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    for (int c = 0; c < width; c++)
                    {
                        using (ComBag cellBag = new ComBag())
                        {
                            dynamic cell = cellBag.Own(sheet.Range[ComFormulaHost.Address(top, left + c)]);
                            if (ComFormulaHost.Boolean((object)cell.MergeCells, "MergeCells")) throw new InvalidOperationException("表头不能合并。");
                            headers[c] = Convert.ToString((object)cell.Value2, CultureInfo.InvariantCulture);
                            if (string.IsNullOrWhiteSpace(headers[c]) || !unique.Add(headers[c])) throw new InvalidOperationException("数据表第一行需要非空且不重复的列名。");
                        }
                    }
                    List<FilterCriterion> conditions = new List<FilterCriterion>();
                    for (int a = 1; a <= Convert.ToInt32(areas.Count); a++)
                    {
                        using (ComBag areaBag = new ComBag())
                        {
                            dynamic area = areaBag.Own(areas.Item(a)), rr = areaBag.Own(area.Rows), cc = areaBag.Own(area.Columns);
                            int ar = Convert.ToInt32(area.Row), ac = Convert.ToInt32(area.Column), rh = Convert.ToInt32(rr.Count), cw = Convert.ToInt32(cc.Count);
                            if (ar <= top || ar + rh > top + height || ac < left || ac + cw > left + width) throw new InvalidOperationException("条件必须位于同一连续数据表的数据行内，不能含表头或跨表格。");
                            if ((long)conditions.Count + (long)rh * cw > 100000) throw new InvalidOperationException("请选择需要的条件行，条件选区不能超过十万个单元格。");
                            for (int r = ar; r < ar + rh; r++) for (int c = ac; c < ac + cw; c++)
                                using (ComBag one = new ComBag()) { dynamic cell = one.Own(sheet.Range[ComFormulaHost.Address(r, c)]); if (ComFormulaHost.Boolean((object)cell.MergeCells, "MergeCells")) throw new InvalidOperationException("条件单元格不能合并。"); object v = cell.Value2; if (v is ErrorWrapper) throw new InvalidOperationException("条件不能包含错误值。"); conditions.Add(new FilterCriterion { Row = r, Column = c - left, Value = v }); }
                        }
                    }
                    object[,] criteria = QuickFilterCriteria.Build(headers, conditions);
                    dynamic book = bag.Own(sheet.Parent); object output = null, temp = null; bool completed = false;
                    using (HostQuiet quiet = new HostQuiet(app))
                    {
                        try
                        {
                            output = WorkbookUi.NewSheet((object)book, "高筛结果", bag, false);
                            temp = WorkbookUi.NewSheet((object)book, "_EI高筛条件", bag, false);
                            dynamic criterionRange = bag.Own(((dynamic)temp).Range["A1:" + ComFormulaHost.Address(criteria.GetLength(0), criteria.GetLength(1))]);
                            criterionRange.NumberFormat = "@"; criterionRange.Value2 = criteria;
                            dynamic destination = bag.Own(((dynamic)output).Range["A1:" + ComFormulaHost.Address(1, width)]);
                            object[,] head = new object[1, width]; for (int c = 0; c < width; c++) head[0, c] = headers[c]; destination.Value2 = head;
                            ((dynamic)output).Activate();
                            // WPS/Excel's native engine defines comparisons, dates,
                            // wildcard escaping and text criteria semantics.
                            source.AdvancedFilter(2, criterionRange, destination, false);
                            dynamic used = bag.Own(((dynamic)output).UsedRange), rows = bag.Own(used.Rows);
                            int count = Math.Max(0, Convert.ToInt32(rows.Count) - 1);
                            object values = used.Value2; used.NumberFormat = "@"; used.Value2 = values;
                            for (int c = 0; c < width; c++) using (ComBag one = new ComBag())
                            {
                                dynamic original = one.Own(sheet.Range[ComFormulaHost.Address(top + 1, left + c)]);
                                dynamic target = one.Own(((dynamic)output).Range[ComFormulaHost.Address(2, c + 1) + ":" + ComFormulaHost.Address(Math.Max(2, count + 1), c + 1)]);
                                target.NumberFormat = original.NumberFormat;
                            }
                            WorkbookUi.StyleTable(output, width, count + 1);
                            WorkbookUi.FitOutput(output, 1, width, count + 1);
                            WorkbookUi.DeleteSheet(app, temp); temp = null; completed = true; return count;
                        }
                        finally
                        {
                            List<string> failures = new List<string>();
                            if (temp != null) try { WorkbookUi.DeleteSheet(app, temp); } catch (Exception ex) { failures.Add("临时条件表清理失败：" + ex.Message); }
                            if (!completed && output != null) try { WorkbookUi.DeleteSheet(app, output); } catch (Exception ex) { failures.Add("未完成结果表清理失败：" + ex.Message); }
                            if (!completed) try { sheet.Activate(); selection.Select(); } catch (Exception ex) { failures.Add("选区恢复失败：" + ex.Message); }
                            if (failures.Count > 0) throw new InvalidOperationException(string.Join("；", failures));
                        }
                    }
                }
            }
        }
    }

    [ComVisible(true), Guid("00024413-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDirectoryAppEvents
    {
        [DispId(1854)] void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sheet, [In, MarshalAs(UnmanagedType.IDispatch)] object hyperlink);
    }
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public sealed class DirectoryEventSink : IDirectoryAppEvents
    {
        internal object App;
        private bool handling;
        public void SheetFollowHyperlink(object sheet, object hyperlink)
        {
            if (handling || App == null) return;
            bool ours = false; handling = true;
            try
            {
                using (ComBag bag = new ComBag())
                {
                    dynamic book = bag.Own(((dynamic)sheet).Parent), names = bag.Own(book.Names);
                    dynamic clicked = bag.Own(((dynamic)hyperlink).Range);
                    for (int i = 1; i <= Convert.ToInt32(names.Count); i++)
                    {
                        using (ComBag one = new ComBag())
                        {
                            dynamic name = one.Own(names.Item(i)); string key = (string)name.Name;
                            int prefix = key.IndexOf("_EI_Dir_A_", StringComparison.Ordinal);
                            if (prefix < 0) continue;
                            dynamic anchor, anchorSheet;
                            // Deleted older directory sheets leave orphan names;
                            // these must not block a newer directory's hyperlinks.
                            try { anchor = one.Own(name.RefersToRange); anchorSheet = one.Own(anchor.Worksheet); } catch { continue; }
                            if (!WorksheetIdentity.SameObject((object)anchorSheet, sheet) || (int)anchor.Row != (int)clicked.Row || (int)anchor.Column != (int)clicked.Column) continue;
                            ours = true;
                            dynamic targetName = one.Own(names.Item(key.Substring(prefix).Replace("_EI_Dir_A_", "_EI_Dir_T_")));
                            dynamic target = one.Own(targetName.RefersToRange), targetSheet = one.Own(target.Worksheet);
                            if (ComFormulaHost.Boolean((object)book.ProtectStructure, "ProtectStructure") && Convert.ToInt32(targetSheet.Visible) != -1)
                                throw new InvalidOperationException("工作簿结构受保护，不能显示隐藏工作表；请先取消结构保护。");
                            int originalVisibility = Convert.ToInt32(targetSheet.Visible);
                            try { targetSheet.Visible = -1; targetSheet.Activate(); ((dynamic)App).Goto(target, false); }
                            catch { if (originalVisibility != -1) { ((dynamic)sheet).Activate(); targetSheet.Visible = originalVisibility; } throw; }
                            return;
                        }
                    }
                }
            }
            catch (Exception ex) { Connect.Log("目录跳转：" + ex); if (ours) UiDialog.Show(null, "无法跳转：" + ex.Message + "\r\n目标已删除时请重新生成目录。", "生成目录", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { handling = false; }
        }
    }
    internal sealed class DirectoryNavigation : IDisposable
    {
        private IConnectionPoint point;
        private DirectoryEventSink sink;
        private int cookie;
        internal DirectoryNavigation(object app)
        {
            IConnectionPointContainer container = app as IConnectionPointContainer;
            if (container == null) throw new InvalidOperationException("当前宿主不支持目录超链接事件。");
            Guid id = new Guid("00024413-0000-0000-C000-000000000046");
            try { container.FindConnectionPoint(ref id, out point); sink = new DirectoryEventSink { App = app }; point.Advise(sink, out cookie); }
            catch { Dispose(); throw; }
        }
        public void Dispose()
        {
            if (sink != null) sink.App = null;
            try { if (point != null && cookie != 0) point.Unadvise(cookie); }
            finally { cookie = 0; if (point != null && Marshal.IsComObject(point)) Marshal.ReleaseComObject(point); point = null; sink = null; }
        }
        internal static int Generate(object app)
        {
            using (ComBag bag = new ComBag())
            {
                dynamic book = bag.Own(((dynamic)app).ActiveWorkbook);
                if (book == null) throw new InvalidOperationException("请先打开工作簿。");
                dynamic sheets = bag.Own(book.Worksheets), names = bag.Own(book.Names);
                List<object> targets = new List<object>();
                for (int i = 1; i <= Convert.ToInt32(sheets.Count); i++) targets.Add(bag.Own(sheets.Item(i)));
                object output = null; List<object> createdNames = new List<object>(); string token = Guid.NewGuid().ToString("N"); bool success = false;
                using (HostQuiet quiet = new HostQuiet(app)) try
                {
                    output = WorkbookUi.NewSheet((object)book, "目录", bag, true);
                    object[,] values = new object[targets.Count + 1, 3]; values[0, 0] = "工作表名称"; values[0, 1] = "跳转"; values[0, 2] = "生成时状态";
                    for (int i = 0; i < targets.Count; i++) { dynamic s = targets[i]; values[i + 1, 0] = s.Name; values[i + 1, 1] = "跳转至本表"; int v = Convert.ToInt32(s.Visible); values[i + 1, 2] = v == -1 ? "可见" : v == 2 ? "非常隐藏" : "隐藏"; }
                    dynamic range = bag.Own(((dynamic)output).Range["A1:C" + (targets.Count + 1)]); range.NumberFormat = "@"; range.Value2 = values;
                    dynamic links = bag.Own(((dynamic)output).Hyperlinks);
                    for (int i = 0; i < targets.Count; i++)
                    {
                        dynamic s = targets[i]; string index = token + "_" + i, anchorAddress = "$B$" + (i + 2);
                        string destination = WorkbookUi.Quote((string)s.Name) + "!$A$1";
                        object targetName = bag.Own(names.Add("_EI_Dir_T_" + index, "=" + destination, false)); createdNames.Add(targetName);
                        object anchorName = bag.Own(names.Add("_EI_Dir_A_" + index, "=" + WorkbookUi.Quote((string)((dynamic)output).Name) + "!" + anchorAddress, false)); createdNames.Add(anchorName);
                        dynamic cell = bag.Own(((dynamic)output).Range[anchorAddress]);
                        // A valid self target triggers the event without the host's
                        // invalid-hidden-target warning; the sink navigates afterward.
                        string address = Convert.ToInt32(s.Visible) == -1 ? destination : WorkbookUi.Quote((string)((dynamic)output).Name) + "!" + anchorAddress;
                        bag.Own(links.Add(cell, "", address, "跳转至 " + (string)s.Name, "跳转至本表"));
                    }
                    WorkbookUi.StyleTable(output, 3, targets.Count + 1);
                    // Apply after Hyperlinks.Add and the shared table style so every
                    // directory cell, including clickable links, retains this format.
                    dynamic directoryFont = bag.Own(range.Font);
                    directoryFont.Bold = true;
                    range.HorizontalAlignment = -4108; // xlCenter
                    range.VerticalAlignment = -4108;
                    WorkbookUi.FitOutput(output, 1, 3, targets.Count + 1);
                    ((dynamic)output).Activate(); success = true; return targets.Count;
                }
                finally
                {
                    if (!success)
                    {
                        List<string> failures = new List<string>();
                        foreach (object name in createdNames) try { ((dynamic)name).Delete(); } catch (Exception ex) { failures.Add(ex.Message); }
                        if (output != null) try { WorkbookUi.DeleteSheet(app, output); } catch (Exception ex) { failures.Add(ex.Message); }
                        if (failures.Count > 0) throw new InvalidOperationException("目录清理不完整：" + string.Join("；", failures));
                    }
                }
            }
        }
    }
}
