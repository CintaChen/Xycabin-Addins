using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // All methods are called on the Ribbon/WinForms host STA. COM never enters a worker.
    internal sealed class SummaryAreaHost : IDisposable
    {
        private readonly object app;
        private readonly Form form;
        private readonly SummaryAreaView view;
        private readonly Action<bool> hostBusy;
        private readonly Func<bool> approveOverwrite;
        private readonly ComBag references = new ComBag();
        private readonly Timer timer = new Timer { Interval = 40 };
        private object book;
        private string bookName, templateSheet, pasteSheet, leadSheet, backup;
        private readonly List<SummaryAreaPlan> areas = new List<SummaryAreaPlan>();
        private SummaryAreaPlan paste, lead;
        private readonly List<SummaryAreaCell> cells = new List<SummaryAreaCell>();
        private List<SummaryAreaPlan> targets = new List<SummaryAreaPlan>();
        private readonly List<SummaryCompany> companies = new List<SummaryCompany>();
        private readonly List<SummaryAreaCell> originalTargetFormulas = new List<SummaryAreaCell>();
        private string stamp, selectedKey;
        private SourceChangeWatch watch;
        private bool ready, running, pending, ticking, loading, failed, handling;
        private int step, formulaCount, touched;
        internal SummaryAreaHost(object app, Form form, SummaryAreaView view, Action<bool> hostBusy = null, Func<bool> approveOverwrite = null)
        {
            this.app = app; this.form = form; this.view = view; this.hostBusy = hostBusy ?? delegate { };
            this.approveOverwrite = approveOverwrite ?? delegate { return UiDialog.Show(form, "将覆盖预检日志中列出的全部目标 Area（包括所选空白格）。中间未选区域不变。确认执行？", "确认覆盖", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.OK; };
            view.Pick += Pick; view.PreviewRequested += delegate { Safe(Preview); };
            view.ExecuteRequested += delegate { Safe(Start); }; view.UndoRequested += delegate { Safe(Undo); };
            view.ConfirmRequested += delegate { Safe(Confirm); }; view.Changed += Invalidate;
            form.Activated += delegate { if (ready) Safe(delegate { CheckBook(); if (watch != null) watch.Validate(); }); };
            form.FormClosing += delegate(object sender, FormClosingEventArgs e) { if (handling || running || pending) { e.Cancel = true; view.Message("请先完成当前操作、确认或撤销本批；备份表：" + backup); } };
            timer.Tick += Tick;
        }
        private void Safe(Action action)
        {
            if (handling) return;
            handling = true; hostBusy(true); view.SetState(true, pending, false);
            try { action(); }
            catch (Exception error) { ready = false; view.Message("失败：" + error.Message); view.SetState(running, pending, false); }
            finally { handling = false; hostBusy(running); view.SetState(running, pending, ready); }
        }
        private void Invalidate() { if (loading) return; ready = false; if (!running) view.SetState(false, pending, false); }
        private void CheckBook()
        {
            if (book == null) throw new InvalidOperationException("请先选择模板区域。");
            using (var bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)app).Workbooks); bool found = false;
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++) if (WorksheetIdentity.SameObject(book, bag.Own(books.Item(i)))) found = true;
                if (!found || Convert.ToString(((dynamic)book).FullName) != bookName) throw new InvalidOperationException("绑定工作簿已关闭或另存为，请关闭本窗口后重选。");
            }
        }
        private object Sheet(string name, ComBag bag) { dynamic sheets = bag.Own(((dynamic)book).Worksheets); return bag.Own(sheets.Item(name)); }
        private static SummaryAreaPlan Coordinates(dynamic range, ComBag bag)
        {
            return new SummaryAreaPlan { Row = Convert.ToInt32(range.Row), Column = Convert.ToInt32(range.Column), Rows = Convert.ToInt32(((dynamic)bag.Own(range.Rows)).Count), Columns = Convert.ToInt32(((dynamic)bag.Own(range.Columns)).Count), Address = Convert.ToString(range.Address) };
        }
        private void Pick(string kind, bool native)
        {
            if (running || pending) return;
            Safe(delegate {
                using (var bag = new ComBag())
                {
                    object chosen = null;
                    try
                    {
                        if (native) { form.Hide(); chosen = DataMigrationHost.PickNativeRange(app, kind == "template" ? "选择模板区域，可按 Ctrl 多选；若原生框不支持，请取消后使用读取当前选区。" : kind == "lead" ? "选择公司编码、全称相邻两列的数据行，不含表头。" : "选择粘贴起点，一个单元格。", "单体汇总批量链接", Type.Missing); }
                        else chosen = ((dynamic)app).Selection;
                        if (chosen == null || chosen is bool) return;
                        dynamic range = bag.Own(chosen), sheet = bag.Own(range.Parent);
                        object selectedBook = sheet.Parent;
                        if (book == null) { book = references.Own(selectedBook); bookName = Convert.ToString(((dynamic)book).FullName); }
                        else { bag.Own(selectedBook); if (!WorksheetIdentity.SameObject(book, selectedBook)) throw new InvalidOperationException("三组选区必须来自同一个工作簿。"); }
                        CheckBook(); string sheetName = Convert.ToString(sheet.Name);
                        dynamic selectedAreas = bag.Own(range.Areas); int count = Convert.ToInt32(selectedAreas.Count);
                        if (kind != "template" && count != 1) throw new InvalidOperationException("此参数只接受一个连续区域。");
                        var values = new List<SummaryAreaPlan>();
                        for (int i = 1; i <= count; i++)
                        {
                            using (var one = new ComBag())
                            {
                                dynamic area = one.Own(selectedAreas.Item(i)), parent = one.Own(area.Parent);
                                if (!WorksheetIdentity.SameObject(parent, sheet)) throw new InvalidOperationException("不接受跨工作表多选。");
                                var value = Coordinates(area, one); value.Sheet = sheetName;
                                if ((long)value.Rows * value.Columns > SummaryAreaEngine.CellLimit) throw new InvalidOperationException("选区过大，请选择实际数据范围。");
                                values.Add(value);
                            }
                        }
                        if (kind == "paste" && (values[0].Rows != 1 || values[0].Columns != 1)) throw new InvalidOperationException("粘贴起点必须是一个单元格。");
                        if (kind == "lead" && values[0].Columns != 2) throw new InvalidOperationException("公司字段必须是相邻两列，第一列编码简称、第二列全称。");
                        if (kind == "template") { int minRow = values.Min(a => a.Row), minColumn = values.Min(a => a.Column); foreach (var a in values) { a.RowOffset = a.Row - minRow; a.ColumnOffset = a.Column - minColumn; } areas.Clear(); areas.AddRange(values); templateSheet = sheetName; }
                        else if (kind == "paste") { paste = values[0]; pasteSheet = sheetName; }
                        else { lead = values[0]; leadSheet = sheetName; loading = true; try { view.Companies.Rows.Clear(); } finally { loading = false; } }
                        view.Address(kind, "[" + ((dynamic)book).Name + "]" + sheetName + "!" + string.Join("; ", values.Select(a => a.Address).ToArray()));
                    }
                    finally { if (native && !form.IsDisposed) { form.Show(); form.Activate(); } }
                }
            });
        }
        private static bool Bool(object value) { return value != null && value != DBNull.Value && Convert.ToBoolean(value); }
        private void RejectSpecial(dynamic range, string address, ComBag bag)
        {
            object merged = range.MergeCells;
            if (merged == null || merged == DBNull.Value || Bool(merged)) throw new InvalidOperationException(address + " 含合并或混合合并区域；MVP 请先取消合并。");
            object array = range.HasArray;
            if (array == DBNull.Value || Bool(array)) throw new InvalidOperationException(address + " 含数组或混合数组公式。");
            try { object spill = range.HasSpill; if (spill == DBNull.Value || Bool(spill)) throw new InvalidOperationException(address + " 含溢出或混合溢出公式。"); }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            catch (System.Runtime.InteropServices.COMException error) { if (error.HResult != unchecked((int)0x80020003) && error.HResult != unchecked((int)0x80020006)) throw; }
        }
        private void ReadCompanies(ComBag bag)
        {
            var selected = new Dictionary<int, bool>();
            foreach (DataGridViewRow row in view.Companies.Rows) selected[Convert.ToInt32(row.Cells["row"].Value)] = Bool(row.Cells["selected"].Value);
            companies.Clear(); dynamic sheet = Sheet(leadSheet, bag); dynamic range = bag.Own(sheet.Range[SummaryAreaEngine.Address(lead)]); RejectSpecial(range, lead.Address, bag);
            object values = range.Value2;
            loading = true;
            try
            {
                view.Companies.Rows.Clear();
                for (int r = 1; r <= lead.Rows; r++)
                {
                    string code = Convert.ToString(SummaryArrays.Pick(values, r - 1, lead.Rows, 0, 2), CultureInfo.InvariantCulture);
                    string name = Convert.ToString(SummaryArrays.Pick(values, r - 1, lead.Rows, 1, 2), CultureInfo.InvariantCulture);
                    if (code.Length == 0 && name.Length == 0) continue;
                    if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(name)) throw new InvalidOperationException("Lead 第 " + (lead.Row + r - 1) + " 行两列必须同时填写。");
                    var company = new SummaryCompany { Row = lead.Row + r - 1, Code = code, Name = name, Selected = !selected.ContainsKey(lead.Row + r - 1) || selected[lead.Row + r - 1] };
                    companies.Add(company); view.Companies.Rows.Add(company.Selected, company.Row, code, name);
                }
            }
            finally { loading = false; }
        }
        private void Preview()
        {
            if (running || pending) return; ready = false; CheckBook();
            if (areas.Count == 0 || paste == null || lead == null) throw new InvalidOperationException("请完成三组选区。");
            if (templateSheet != pasteSheet) throw new InvalidOperationException("模板和粘贴起点必须属于同一工作表。");
            if (!Directory.Exists(view.Folder.Text)) throw new InvalidOperationException("项目文件夹不存在。");
            using (var bag = new ComBag())
            {
                if (Bool(((dynamic)book).ReadOnly) || Bool(((dynamic)book).ProtectStructure)) throw new InvalidOperationException("工作簿只读或结构受保护，不能执行及备份。");
                dynamic sheet = Sheet(templateSheet, bag);
                if (Bool(sheet.ProtectContents)) throw new InvalidOperationException("目标工作表受保护。");
                ReadCompanies(bag); var picked = companies.Where(c => c.Selected).ToList();
                targets = SummaryAreaEngine.Targets(areas, paste.Row, paste.Column, picked.Count);
                if (leadSheet == templateSheet && targets.Any(t => SummaryAreaEngine.Intersects(t, lead))) throw new InvalidOperationException("目标不能覆盖 Lead 公司字段。");
                cells.Clear(); var errors = new List<string>();
                foreach (var area in areas)
                {
                    dynamic range = bag.Own(sheet.Range[SummaryAreaEngine.Address(area)]); RejectSpecial(range, area.Address, bag);
                    for (int row = area.Row; row < area.Row + area.Rows; row++) for (int column = area.Column; column < area.Column + area.Columns; column++)
                    {
                        using (var one = new ComBag())
                        {
                            dynamic cell = one.Own(sheet.Cells[row, column]); object value = cell.Value2;
                            bool formula = Bool(cell.HasFormula); if (!formula && (value == null || Convert.ToString(value).Length == 0)) continue;
                            string text = Convert.ToString(formula ? cell.Formula : value, CultureInfo.InvariantCulture);
                            try
                            {
                                if (formula) SummaryAreaEngine.NativeFormula(text);
                                else foreach (var company in picked) SummaryTextReference.External(view.Folder.Text, company.Name, text);
                                cells.Add(new SummaryAreaCell { Row = row, Column = column, Text = text, Formula = formula, R1C1 = formula ? Convert.ToString(cell.FormulaR1C1) : null });
                            }
                            catch (Exception error) { errors.Add(CellAddress.Name(row, column) + " | " + text + " | " + error.Message); }
                        }
                    }
                }
                // Validate filenames even for an entirely blank template.
                foreach (var company in picked)
                {
                    try { SummaryTextReference.External(view.Folder.Text, company.Name, "BS!A1"); if (view.CheckFiles.Checked && !File.Exists(Path.Combine(view.Folder.Text, company.Name + ".xlsx"))) throw new InvalidOperationException("来源文件不存在。"); }
                    catch (Exception error) { errors.Add("Lead 第 " + company.Row + " 行 | " + company.Name + " | " + error.Message); }
                }
                for (int i = 0; i < targets.Count; i++)
                {
                    var source = areas[i % areas.Count]; var target = targets[i];
                    foreach (var cell in cells.Where(c => c.Formula && c.Row >= source.Row && c.Row < source.Row + source.Rows && c.Column >= source.Column && c.Column < source.Column + source.Columns))
                        try { SummaryAreaEngine.ValidateShift(cell.Text, target.Row - source.Row, target.Column - source.Column); }
                        catch (Exception error) { errors.Add(CellAddress.Name(cell.Row, cell.Column) + " | " + cell.Text + " | " + error.Message); }
                }
                int occupied = 0; originalTargetFormulas.Clear();
                foreach (var target in targets)
                {
                    dynamic range = bag.Own(sheet.Range[SummaryAreaEngine.Address(target)]); RejectSpecial(range, SummaryAreaEngine.Address(target), bag);
                    object values = range.Formula;
                    for (int r = 1; r <= target.Rows; r++) for (int c = 1; c <= target.Columns; c++)
                    {
                        object prior = SummaryArrays.Pick(values, r - 1, target.Rows, c - 1, target.Columns);
                        if (SummaryArrays.HasContent(prior)) occupied++;
                        if (prior is string && ((string)prior).StartsWith("=", StringComparison.Ordinal))
                        using (var one = new ComBag())
                        {
                            dynamic cell = one.Own(sheet.Cells[target.Row + r - 1, target.Column + c - 1]);
                            if (Bool(cell.HasFormula)) originalTargetFormulas.Add(new SummaryAreaCell { Row = target.Row + r - 1, Column = target.Column + c - 1, Text = (string)prior });
                        }
                    }
                    view.Message("将覆盖 " + templateSheet + "!" + SummaryAreaEngine.Address(target));
                }
                foreach (string error in errors) view.Message("阻断：" + error);
                if (errors.Count > 0) throw new InvalidOperationException("预检有 " + errors.Count + " 个阻断项；未修改任何工作表。");
                stamp = Fingerprint(); selectedKey = SelectionKey();
                if (watch != null) watch.Dispose(); watch = new SourceChangeWatch(app, true); watch.WatchBook(book);
                if (!watch.Available) view.Message("宿主不支持变化事件；执行前将重新读取内容和格式指纹。");
                ready = true; view.SetState(false, false, true);
                view.Message("预检通过：" + picked.Count + " 家公司，" + targets.Count + " 个目标 Area；已有内容 " + occupied + " 格。执行前再次确认覆盖；未保存。");
            }
        }
        private string SelectionKey()
        { return view.Folder.Text + "|" + view.CheckFiles.Checked + "|" + string.Join(",", view.Companies.Rows.Cast<DataGridViewRow>().Select(r => Convert.ToString(r.Cells["row"].Value) + ":" + Convert.ToString(r.Cells["selected"].Value)).ToArray()); }
        private string Fingerprint()
        {
            CheckBook(); var text = new StringBuilder(bookName);
            using (var bag = new ComBag())
            {
                foreach (var part in areas.Concat(targets).Concat(new[] { lead }))
                {
                    dynamic sheet = Sheet(part == lead ? leadSheet : templateSheet, bag);
                    if (Bool(sheet.ProtectContents)) throw new InvalidOperationException("工作表保护状态变化。");
                    for (int r = part.Row; r < part.Row + part.Rows; r++) for (int c = part.Column; c < part.Column + part.Columns; c++)
                    using (var one = new ComBag())
                    {
                        dynamic cell = one.Own(sheet.Cells[r, c]);
                        Append(text, cell.Formula); Append(text, cell.NumberFormat); Append(text, cell.Style); Append(text, cell.MergeCells);
                        dynamic font = one.Own(cell.Font), fill = one.Own(cell.Interior); Append(text, font.Name); Append(text, font.Size); Append(text, font.Bold); Append(text, font.Italic); Append(text, font.Color); Append(text, fill.Color);
                        Append(text, cell.HorizontalAlignment); Append(text, cell.VerticalAlignment); Append(text, cell.WrapText);
                        dynamic comment = one.Own(cell.Comment); Append(text, comment == null ? "" : comment.Text());
                    }
                }
            }
            using (var hash = System.Security.Cryptography.SHA256.Create()) return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(text.ToString())));
        }
        private static void Append(StringBuilder text, object value) { string s = Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""; text.Append(s.Length).Append(':').Append(s); }
        private void Start()
        {
            if (!ready || running || pending) throw new InvalidOperationException("请先重新预检。");
            CheckBook(); if (watch != null) watch.Validate();
            if (selectedKey != SelectionKey() || stamp != Fingerprint()) throw new InvalidOperationException("参数、名单、模板或目标状态变化，请重新预检。");
            if (!approveOverwrite()) return;
            if (selectedKey != SelectionKey() || stamp != Fingerprint()) throw new InvalidOperationException("确认期间工作簿变化，请重新预检。");
            if (watch != null) watch.Validate();
            using (var quiet = new HostQuiet(app)) using (var links = new HostLinkQuiet(app)) using (var bag = new ComBag())
            { backup = SummaryWriter.Backup(app, book, Sheet(templateSheet, bag)); ((dynamic)Sheet(templateSheet, bag)).Activate(); }
            pending = true; running = true; ready = failed = false; step = formulaCount = touched = 0; view.SetState(true, true, false);
            view.Message("已建立原始模板及目标备份：" + backup + "。逐 Area 执行，不复制外接矩形。"); timer.Start();
        }
        private void Tick(object sender, EventArgs e)
        {
            if (ticking || !running) return; ticking = true; timer.Stop();
            try
            {
                CheckBook(); if (watch != null) watch.Validate();
                var picked = companies.Where(c => c.Selected).ToList();
                using (var quiet = new HostQuiet(app)) using (var links = new HostLinkQuiet(app)) using (var bag = new ComBag())
                {
                    dynamic sheet = Sheet(templateSheet, bag), original = Sheet(backup, bag);
                    var sourceArea = areas[step % areas.Count]; var targetArea = targets[step]; var company = picked[step / areas.Count];
                    dynamic source = bag.Own(original.Range[SummaryAreaEngine.Address(sourceArea)]), target = bag.Own(sheet.Range[SummaryAreaEngine.Address(targetArea)]);
                    touched = step + 1; source.Copy(target);
                    foreach (var cell in cells.Where(c => c.Row >= sourceArea.Row && c.Row < sourceArea.Row + sourceArea.Rows && c.Column >= sourceArea.Column && c.Column < sourceArea.Column + sourceArea.Columns))
                    {
                        using (var one = new ComBag())
                        {
                            dynamic destination = one.Own(sheet.Cells[targetArea.Row + cell.Row - sourceArea.Row, targetArea.Column + cell.Column - sourceArea.Column]);
                            if (!cell.Formula)
                            {
                                string formula = SummaryTextReference.External(view.Folder.Text, company.Name, cell.Text); destination.Formula = formula;
                                if (!Bool(destination.HasFormula) || !Equivalent(Convert.ToString(destination.Formula), formula)) throw new InvalidOperationException("真实公式写入核验失败：" + destination.Address);
                            }
                            else if (!Bool(destination.HasFormula) || !Equivalent(Convert.ToString(destination.FormulaR1C1), cell.R1C1)) throw new InvalidOperationException("原生公式相对引用复制核验失败：" + destination.Address);
                            formulaCount++;
                        }
                    }
                    step++; view.Message("已完成 " + step + "/" + targets.Count + "：" + company.Name + " → " + SummaryAreaEngine.Address(targetArea));
                    if (step == targets.Count)
                    {
                        int errors = 0;
                        foreach (var area in targets)
                        {
                            dynamic range = bag.Own(sheet.Range[SummaryAreaEngine.Address(area)]);
                            try { range.Calculate(); } catch (Exception error) { view.Message("重新计算失败：" + error.Message); }
                            for (int r = area.Row; r < area.Row + area.Rows; r++) for (int c = area.Column; c < area.Column + area.Columns; c++)
                            using (var one = new ComBag()) { dynamic value = one.Own(sheet.Cells[r, c]); if (value.Value2 is System.Runtime.InteropServices.ErrorWrapper || Convert.ToString(value.Text).StartsWith("#", StringComparison.Ordinal)) errors++; }
                        }
                        view.Message("公式写入核验成功 " + formulaCount + " 格；已尝试重新计算。数值异常 " + errors + " 格。链接未强制更新：源文件关闭或访问策略可能保留缓存，请在 Excel 检查链接状态。未保存，请确认或撤销。"); running = false; view.SetState(false, true, false);
                    }
                }
            }
            catch (Exception error)
            {
                running = false; failed = true; view.Message("执行失败：" + error.Message + "；正在撤销已涉及 Area。");
                try { Restore(touched); pending = false; view.Message("已还原至执行前状态，备份保留：" + backup); }
                catch (Exception restore) { pending = true; view.Message("自动还原失败：" + restore.Message + "；备份保留，请重试撤销。禁止确认丢弃恢复信息。"); }
                ready = false; view.SetState(false, pending, false);
            }
            finally { ticking = false; hostBusy(running); if (running) timer.Start(); }
        }
        private static bool Equivalent(string a, string b) { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }
        private void Restore(int count)
        {
            CheckBook(); using (var quiet = new HostQuiet(app)) using (var links = new HostLinkQuiet(app)) using (var bag = new ComBag())
            {
                dynamic sheet = Sheet(templateSheet, bag), original = Sheet(backup, bag);
                for (int i = 0; i < count; i++)
                {
                    string address = SummaryAreaEngine.Address(targets[i]); dynamic source = bag.Own(original.Range[address]), target = bag.Own(sheet.Range[address]);
                    target.Clear(); source.Copy(target);
                    var area = targets[i];
                    foreach (var prior in originalTargetFormulas)
                    {
                        // Worksheet.Copy may redirect explicit self-sheet references to
                        // the backup. Restore actual formula text from the read-only plan.
                        if (prior.Row < area.Row || prior.Row >= area.Row + area.Rows || prior.Column < area.Column || prior.Column >= area.Column + area.Columns) continue;
                        using (var one = new ComBag())
                        {
                            dynamic cell = one.Own(sheet.Cells[prior.Row, prior.Column]); cell.Formula = prior.Text;
                            if (!string.Equals(Convert.ToString(cell.Formula), prior.Text, StringComparison.Ordinal)) throw new InvalidOperationException("还原公式读回失败：" + CellAddress.Name(prior.Row, prior.Column));
                        }
                    }
                }
            }
        }
        private void Undo()
        { if (!pending || running) return; Restore(touched); pending = ready = false; view.SetState(false, false, false); view.Message("已撤销全部受影响目标 Area，未选中间区域未修改；备份保留，未保存。"); }
        private void Confirm()
        {
            if (!pending || running) return;
            if (failed || step != targets.Count) throw new InvalidOperationException("失败批次只能撤销，不能确认。");
            pending = ready = false; view.SetState(false, false, false); view.Message("已确认并结束本批回滚会话；未保存工作簿。备份仍保留：" + backup);
        }
        public void Dispose() { timer.Stop(); timer.Dispose(); hostBusy(false); if (watch != null) watch.Dispose(); references.Dispose(); }
    }
}
