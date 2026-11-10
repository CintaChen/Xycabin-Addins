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
        private HostCalculation calculation;
        private SummaryWorkbookBinding presetBinding;
        private readonly System.Diagnostics.Stopwatch executionClock = new System.Diagnostics.Stopwatch();
        private int bulkWrites;
        private SummaryAreaPlan companyHeader, yearSource, yearTarget;
        private readonly List<SummaryAreaPlan> headerTargets = new List<SummaryAreaPlan>();
        private readonly List<string> headerValues = new List<string>();
        private bool headersTouched;

        private readonly ComBag references = new ComBag();
        private readonly Timer timer = new Timer { Interval = 40 };
        private object book;
        private string bookName, templateSheet, pasteSheet, leadSheet, backup;
        private readonly List<SummaryAreaPlan> areas = new List<SummaryAreaPlan>();
        private SummaryAreaPlan paste, lead;
        private readonly List<SummaryAreaPlan> leadAreas = new List<SummaryAreaPlan>();
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
            this.approveOverwrite = approveOverwrite ?? delegate { return UiDialog.Show(form, "将覆盖预检日志中列出的全部目标 Area 和启用的名称、年度表头（包括所选空白格）。中间未选区域不变。确认执行？", "确认覆盖", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.OK; };
            view.AttachPresets(new BatchPresetActions(form, "summary", CapturePreset, ApplyPreset, delegate { return !handling && !running && !pending; }, view.Message, ContextAvailable));
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
        private bool ContextAvailable(string path)
        {
            using (var bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)app).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++) if (string.Equals(Convert.ToString(((dynamic)bag.Own(books.Item(i))).FullName), path, StringComparison.OrdinalIgnoreCase)) return true;
                return false;
            }
        }
        internal BatchPreset CapturePreset()
        {
            var preset = new BatchPreset { Feature = "summary", Name = "区域链接", Context = book == null ? "" : bookName, Values = new Dictionary<string, string> {
                { "mode", "area-mvp" }, { "templateSheet", templateSheet ?? "" }, { "templateAreas", SummaryAreaSettings.Encode(areas) },
                { "pasteSheet", pasteSheet ?? "" }, { "pasteArea", paste == null ? "" : SummaryAreaSettings.Encode(new[] { paste }) },
                { "leadSheet", leadSheet ?? "" }, { "leadAreas", SummaryAreaSettings.Encode(LeadSelections()) }, { "projectFolder", view.Folder.Text }, { "checkFiles", view.CheckFiles.Checked.ToString() }
            } };
            preset.Values["headersEnabled"] = view.HeadersEnabled.Checked.ToString();
            var names = new[] { "companyHeader", "yearSource", "yearTarget" }; var parts = new[] { companyHeader, yearSource, yearTarget };
            for(int i=0;i<names.Length;i++) { preset.Values[names[i]] = parts[i] == null ? "" : SummaryAreaSettings.Encode(new[] { parts[i] }); preset.Values[names[i]+"Sheet"] = parts[i] == null ? "" : parts[i].Sheet; }
            return preset;
        }
        internal void ApplyPreset(BatchPreset preset)
        {
            if (handling || running || pending) throw new InvalidOperationException("执行中或待确认时不能应用预设。");
            var settings = SummaryAreaSettings.Read(preset); var candidate = new SummaryWorkbookBinding(app); bool committed = false;
            handling = true; hostBusy(true); view.SetState(true, false, false);
            try
            {
                if (ContextAvailable(preset.Context)) candidate.Choose(preset.Context); else candidate.Browse(preset.Context);
                object selectedBook = candidate.RequireBook();
                using (var bag = new ComBag())
                {
                    dynamic sheets = bag.Own(((dynamic)selectedBook).Worksheets);
                    foreach (var part in settings.Template.Concat(settings.Lead).Concat(new[] { settings.Paste }))
                    { dynamic sheet = bag.Own(sheets.Item(part.Sheet)); dynamic range = bag.Own(sheet.Range[SummaryAreaEngine.Address(part)]); RejectSpecial(range, part.Address, bag); }
                }
                using (var bag = new ComBag())
                {
                    dynamic sheets = bag.Own(((dynamic)selectedBook).Worksheets);
                    foreach(var part in new[]{settings.CompanyHeader,settings.YearSource,settings.YearTarget}.Where(p=>p!=null))
                    { dynamic sheet = bag.Own(sheets.Item(part.Sheet)); dynamic range = bag.Own(sheet.Range[SummaryAreaEngine.Address(part)]); if(part==settings.CompanyHeader) RejectArraySpill(range,"公司名称",bag); else RejectSpecial(range,"表头选区",bag); }
                }
                var oldBinding = presetBinding;
                book = selectedBook; bookName = Convert.ToString(((dynamic)book).FullName); presetBinding = candidate;
                templateSheet = settings.Template[0].Sheet; pasteSheet = settings.Paste.Sheet; leadSheet = settings.Lead[0].Sheet;
                areas.Clear(); areas.AddRange(settings.Template); paste = settings.Paste; leadAreas.Clear(); leadAreas.AddRange(settings.Lead); lead = leadAreas[0];
                loading = true;
                try
                {
                    companyHeader=settings.CompanyHeader; yearSource=settings.YearSource; yearTarget=settings.YearTarget; view.HeadersEnabled.Checked=settings.HeadersEnabled;
                    var headerNames=new[]{"companyHeader","yearSource","yearTarget"}; var headerParts=new[]{companyHeader,yearSource,yearTarget};
                    for(int i=0;i<headerNames.Length;i++) view.Address(headerNames[i],headerParts[i]==null ? "未选择" : headerParts[i].Sheet+"!"+SummaryAreaEngine.Address(headerParts[i]));
                    view.HeaderPreview.Text="请重新预检以生成映射预览。";
                    view.Folder.Text = settings.Folder; view.CheckFiles.Checked = settings.CheckFiles; view.Companies.Rows.Clear(); view.ReferenceSources.Rows.Clear();
                    view.Address("template", templateSheet + "!" + string.Join("; ", areas.Select(a => a.Address).ToArray()));
                    view.Address("paste", pasteSheet + "!" + paste.Address); view.Address("lead", leadSheet + "!" + string.Join("; ", leadAreas.Select(a => a.Address).ToArray()));
                }
                finally { loading = false; }
                ready = false; companies.Clear(); cells.Clear(); targets.Clear(); committed = true;
                if (watch != null) { watch.Dispose(); watch = null; }
                if (oldBinding != null) oldBinding.Dispose();
                view.Message("已恢复区域预设；公司名单不从预设读取，请重新读取 Lead 并预检。未修改、未保存工作表。");
            }
            finally
            {
                if (!committed)
                {
                    try { if (candidate.OpenedByBrowse) using (var quiet = new HostQuiet(app)) ((dynamic)candidate.RequireBook()).Close(false); }
                    finally { candidate.Dispose(); }
                }
                handling = false; hostBusy(false); view.SetState(false, false, ready);
            }
        }
        private void CheckBook()
        {
            if (book == null) throw new InvalidOperationException("请先选择已编制区域。");
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
        private List<SummaryAreaPlan> LeadSelections()
        { return leadAreas.Count > 0 ? leadAreas : lead == null ? new List<SummaryAreaPlan>() : new List<SummaryAreaPlan> { lead }; }
        internal bool ActivateLeadSheet()
        {
            using (var bag = new ComBag())
            {
                object context = book;
                if (context == null) context = bag.Own(((dynamic)app).ActiveWorkbook); else CheckBook();
                if (context == null) return false;
                dynamic sheets = bag.Own(((dynamic)context).Worksheets);
                for (int i = 1; i <= Convert.ToInt32(sheets.Count); i++)
                using (var one = new ComBag())
                {
                    dynamic sheet = one.Own(sheets.Item(i));
                    if (!string.Equals(Convert.ToString(sheet.Name), "Lead", StringComparison.OrdinalIgnoreCase)) continue;
                    if (Convert.ToInt32(sheet.Visible) != -1) { view.Message("Lead 表已隐藏，未自动取消隐藏；请自行选择可见的公司字段表。"); return false; }
                    ((dynamic)context).Activate(); sheet.Activate();
                    view.Message("已跳转到 " + sheet.Name + "；可 Ctrl 选择分段数据或不相邻的两列，同一行需同时选中编码和全称。"); return true;
                }
                view.Message("未找到 Lead 表，请在现有工作表中选择公司字段。"); return false;
            }
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
                        if (native) { if (kind == "lead") ActivateLeadSheet(); form.Hide(); chosen = DataMigrationHost.PickNativeRange(app, kind == "template" ? "选择已编制区域，可按 Ctrl 多选；若原生框不支持，请取消后使用读取当前选区。" : kind == "lead" ? "选择编码和全称两列的数据行，可 Ctrl 多选分段或不相邻两列；同一行需配对，不含表头。" : kind == "yearSource" ? "选择单行连续年度来源，数量应等于已编制区域横向跨度；保留显示文字和顺序。" : kind == "companyHeader" ? "选择第一家公司名称目标单元格，支持既有合并区域，不创建或拆分合并。" : kind == "yearTarget" ? "选择第一家公司年度写入起点，一个单元格。" : "选择粘贴起点，一个单元格。", "单体汇总批量链接", Type.Missing); }
                        else chosen = ((dynamic)app).Selection;
                        if (chosen == null || chosen is bool) return;
                        dynamic range = bag.Own(chosen), sheet = bag.Own(range.Parent);
                        object selectedBook = sheet.Parent;
                        if (book == null) { book = references.Own(selectedBook); bookName = Convert.ToString(((dynamic)book).FullName); }
                        else { bag.Own(selectedBook); if (!WorksheetIdentity.SameObject(book, selectedBook)) throw new InvalidOperationException("三组选区必须来自同一个工作簿。"); }
                        CheckBook(); string sheetName = Convert.ToString(sheet.Name);
                        dynamic selectedAreas = bag.Own(range.Areas); int count = Convert.ToInt32(selectedAreas.Count);
                        if (kind != "template" && kind != "lead" && count != 1) throw new InvalidOperationException("此参数只接受一个连续区域。");
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
                        if (kind == "lead") SummaryAreaEngine.LeadRows(values);
                        if(kind=="companyHeader" && (values[0].Rows!=1 || values[0].Columns!=1))
                        {
                            dynamic first=bag.Own(sheet.Range[CellAddress.Name(values[0].Row,values[0].Column)]);
                            if(!Bool((object)first.MergeCells)) throw new InvalidOperationException("请选择一个公司名称单元格或一个既有合并区域。");
                            dynamic merged=bag.Own(first.MergeArea); var bound=Coordinates(merged,bag); var chosenPart=values[0];
                            if(chosenPart.Row<bound.Row||chosenPart.Column<bound.Column||(long)chosenPart.Row+chosenPart.Rows>bound.Row+bound.Rows||(long)chosenPart.Column+chosenPart.Columns>bound.Column+bound.Columns)throw new InvalidOperationException("公司名称选区必须位于同一个既有合并区域内。");
                            values[0]=new SummaryAreaPlan{Sheet=sheetName,Row=bound.Row,Column=bound.Column,Rows=1,Columns=1,Address=CellAddress.Name(bound.Row,bound.Column)};
                        }
                        if (kind == "yearTarget" && (values[0].Rows!=1||values[0].Columns!=1))throw new InvalidOperationException("年度起点请选择一个单元格。");
                        if (kind == "yearSource" && values[0].Rows!=1)throw new InvalidOperationException("年度来源只接受单行连续区域。");
                        if (kind == "companyHeader") companyHeader=values[0];
                        else if (kind == "yearSource") yearSource=values[0];
                        else if (kind == "yearTarget") yearTarget=values[0];
                        else if (kind == "template") { int minRow = values.Min(a => a.Row), minColumn = values.Min(a => a.Column); foreach (var a in values) { a.RowOffset = a.Row - minRow; a.ColumnOffset = a.Column - minColumn; } areas.Clear(); areas.AddRange(values); templateSheet = sheetName; }
                        else if (kind == "paste") { paste = values[0]; pasteSheet = sheetName; }
                        else { lead = values[0]; leadAreas.Clear(); leadAreas.AddRange(values); leadSheet = sheetName; loading = true; try { view.Companies.Rows.Clear(); } finally { loading = false; } }
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
            RejectArraySpill(range,address,bag);
        }
        private void RejectArraySpill(dynamic range,string address,ComBag bag)
        {
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
            var parts = LeadSelections(); var bindings = SummaryAreaEngine.LeadRows(parts);
            companies.Clear(); dynamic sheet = Sheet(leadSheet, bag);
            foreach (var part in parts)
            using (var one = new ComBag()) { dynamic range = one.Own(sheet.Range[SummaryAreaEngine.Address(part)]); RejectSpecial(range, part.Address, one); }
            loading = true;
            try
            {
                view.Companies.Rows.Clear();
                foreach (var binding in bindings)
                using (var one = new ComBag())
                {
                    dynamic codeCell = one.Own(sheet.Cells[binding.Row, binding.CodeColumn]), nameCell = one.Own(sheet.Cells[binding.Row, binding.NameColumn]);
                    object codeValue = codeCell.Value2, nameValue = nameCell.Value2;
                    if (SummaryAreaEngine.IsError(codeValue) || SummaryAreaEngine.IsError(nameValue)) throw new InvalidOperationException("Lead 第 " + binding.Row + " 行含错误值。");
                    string code = Convert.ToString(codeValue, CultureInfo.InvariantCulture), name = Convert.ToString(nameValue, CultureInfo.InvariantCulture);
                    if (code.Length == 0 && name.Length == 0) continue;
                    if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(name)) throw new InvalidOperationException("Lead 第 " + binding.Row + " 行两列必须同时填写。");
                    var company = new SummaryCompany { Row = binding.Row, Code = code, Name = name, Selected = !selected.ContainsKey(binding.Row) || selected[binding.Row] };
                    companies.Add(company); view.Companies.Rows.Add(company.Selected, company.Row, code, name);
                }
            }
            finally { loading = false; }
        }
        private void Preview()
        {
            if (running || pending) return; ready = false; CheckBook();
            if (areas.Count == 0 || paste == null || lead == null) throw new InvalidOperationException("请完成三组选区。");
            if (templateSheet != pasteSheet) throw new InvalidOperationException("已编制区域和粘贴起点必须属于同一工作表。");
            if (!Directory.Exists(view.ProjectFolder)) throw new InvalidOperationException("项目文件夹不存在。");
            using (var bag = new ComBag())
            {
                if (Bool(((dynamic)book).ReadOnly) || Bool(((dynamic)book).ProtectStructure)) throw new InvalidOperationException("工作簿只读或结构受保护，不能执行及备份。");
                dynamic sheet = Sheet(templateSheet, bag);
                if (Bool(sheet.ProtectContents)) throw new InvalidOperationException("目标工作表受保护。");
                ReadCompanies(bag); var picked = companies.Where(c => c.Selected).ToList();
                targets = SummaryAreaEngine.Targets(areas, paste.Row, paste.Column, picked.Count);
                if (leadSheet == templateSheet && targets.Any(t => LeadSelections().Any(part => SummaryAreaEngine.Intersects(t, part)))) throw new InvalidOperationException("目标不能覆盖 Lead 公司字段。");
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
                                SummaryExistingFormula existing=null;bool literal=!formula && (!(value is string)||text.StartsWith("=",StringComparison.Ordinal)||!System.Text.RegularExpressions.Regex.IsMatch(text,@"!\$?[A-Za-z]{1,3}\$?[0-9]"));
                                if (formula) { existing=SummaryExistingFormula.Read(text); if(!existing.Spans.Any(x=>x.Kind!=RefKind.Local))existing=null; }
                                else if(!literal) foreach (var company in picked) SummaryTextReference.External(view.ProjectFolder, company.Name, text);
                                cells.Add(new SummaryAreaCell { Row = row, Column = column, Text = text, Formula = formula, Literal=literal,Existing=existing,R1C1 = formula ? Convert.ToString(cell.FormulaR1C1) : null });
                            }
                            catch (Exception error) { errors.Add(CellAddress.Name(row, column) + " | " + text + " | " + error.Message); }
                        }
                    }
                }
                ReviewReferences(errors);
                // Validate filenames even for an entirely blank template.
                foreach (var company in picked)
                {
                    try { SummaryTextReference.External(view.ProjectFolder, company.Name, "BS!A1"); if (view.CheckFiles.Checked && !File.Exists(Path.Combine(view.ProjectFolder, company.Name + ".xlsx"))) throw new InvalidOperationException("来源文件不存在。"); }
                    catch (Exception error) { errors.Add("Lead 第 " + company.Row + " 行 | " + company.Name + " | " + error.Message); }
                }
                for (int i = 0; i < targets.Count; i++)
                {
                    var source = areas[i % areas.Count]; var target = targets[i];
                    foreach (var cell in cells.Where(c => c.Formula && c.Row >= source.Row && c.Row < source.Row + source.Rows && c.Column >= source.Column && c.Column < source.Column + source.Columns))
                        try { if(cell.Existing==null)SummaryAreaEngine.ValidateShift(cell.Text, target.Row - source.Row, target.Column - source.Column); else cell.Existing.Render(view.ProjectFolder,picked[i/areas.Count].Name,target.Row-source.Row,target.Column-source.Column); }
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
                PrepareHeaders(bag, picked);
                stamp = Fingerprint(); selectedKey = SelectionKey();
                if (watch != null) watch.Dispose(); watch = new SourceChangeWatch(app, true); watch.WatchBook(book);
                if (!watch.Available) view.Message("宿主不支持变化事件；执行前将重新读取内容和格式指纹。");
                ready = true; view.SetState(false, false, true);
                view.Message("预检通过：" + picked.Count + " 家公司，" + targets.Count + " 个目标 Area，表头写入 "+headerTargets.Count+" 项，阻断异常 0；复制区域已有内容 " + occupied + " 格。执行前再次确认覆盖；未保存。");
            }
        }
        private string SelectionKey()
        { return string.Join(";",view.ReferenceSources.Rows.Cast<DataGridViewRow>().Select(r=>Convert.ToString(r.Cells["source"].Value)+":"+Convert.ToString(r.Cells["decision"].Value)).ToArray()) + "|" + view.HeadersEnabled.Checked + "|" + string.Join(";",new[]{companyHeader,yearSource,yearTarget}.Select(p=>p==null ? "" : p.Sheet+"!"+SummaryAreaSettings.Encode(new[]{p})).ToArray()) + "|" + view.Folder.Text + "|" + view.CheckFiles.Checked + "|" + string.Join(",", view.Companies.Rows.Cast<DataGridViewRow>().Select(r => Convert.ToString(r.Cells["row"].Value) + ":" + Convert.ToString(r.Cells["selected"].Value)).ToArray()); }
        private string Fingerprint()
        {
            CheckBook(); var clock = System.Diagnostics.Stopwatch.StartNew(); var text = new StringBuilder(bookName); int xmlAreas = 0, fallbackAreas = 0;
            using (var bag = new ComBag())
            {
                var leadParts = LeadSelections();
                foreach (var part in areas.Concat(targets).Concat(leadParts).Concat(headerTargets).Concat(view.HeadersEnabled.Checked && yearSource!=null ? new[]{yearSource} : new SummaryAreaPlan[0]))
                {
                    dynamic sheet = Sheet(leadParts.Contains(part) ? leadSheet : part.Sheet ?? templateSheet, bag);
                    if (Bool(sheet.ProtectContents)) throw new InvalidOperationException("工作表保护状态变化。");
                    using (var snapshot = new ComBag())
                    {
                        dynamic range = snapshot.Own(sheet.Range[SummaryAreaEngine.Address(part)]); string xml = null;
                        // Company names may themselves be formulas. Their evaluated
                        // values determine filenames and must invalidate the plan.
                        if (leadParts.Contains(part) || (view.HeadersEnabled.Checked && part==yearSource))
                        {
                            object values = range.Value2;
                            for (int r = 0; r < part.Rows; r++) for (int c = 0; c < part.Columns; c++) Append(text, SummaryArrays.Pick(values, r, part.Rows, c, part.Columns));
                        }
                        if(view.HeadersEnabled.Checked && part==yearSource) for(int c=0;c<part.Columns;c++) { dynamic displayed=snapshot.Own(sheet.Range[CellAddress.Name(part.Row,part.Column+c)]); Append(text,displayed.Text); }
                        try { xml = SummaryAreaEngine.XmlState(Convert.ToString(range.Value[11])); }
                        catch (Exception error) { UiDiagnostics.Record("summary-xml-fallback", error); }
                        if (xml != null) { Append(text, "xml"); Append(text, xml); AppendComments(sheet, part, text); xmlAreas++; continue; }
                    }
                    Append(text, "cells"); fallbackAreas++;
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
            view.Message("内容与格式状态校验耗时 " + clock.Elapsed.TotalSeconds.ToString("F2", CultureInfo.InvariantCulture) + " 秒（整块 " + xmlAreas + "，兼容回退 " + fallbackAreas + "）。");
            using (var hash = System.Security.Cryptography.SHA256.Create()) return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(text.ToString())));
        }
        private static void AppendComments(dynamic sheet, SummaryAreaPlan part, StringBuilder text)
        {
            using (var bag = new ComBag())
            {
                dynamic comments = bag.Own(sheet.Comments); var entries = new List<string>();
                for (int i = 1; i <= Convert.ToInt32(comments.Count); i++)
                using (var one = new ComBag())
                {
                    dynamic comment = one.Own(comments.Item(i)), cell = one.Own(comment.Parent);
                    int r = Convert.ToInt32(cell.Row), c = Convert.ToInt32(cell.Column);
                    if (r >= part.Row && r < part.Row + part.Rows && c >= part.Column && c < part.Column + part.Columns)
                        entries.Add(CellAddress.Name(r, c) + "|" + Convert.ToString(comment.Text()));
                }
                entries.Sort(StringComparer.Ordinal); foreach (string entry in entries) Append(text, entry);
            }
        }
        private static void Append(StringBuilder text, object value) { string s = Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""; text.Append(s.Length).Append(':').Append(s); }
        private void Start()
        {
            if (!ready || running || pending) throw new InvalidOperationException("请先重新预检。");
            CheckBook(); if (watch != null) watch.Validate();
            if (selectedKey != SelectionKey()) throw new InvalidOperationException("参数或名单变化，请重新预检。");
            if (!approveOverwrite()) return;
            if (selectedKey != SelectionKey() || stamp != Fingerprint()) throw new InvalidOperationException("确认期间工作簿变化，请重新预检。");
            if (watch != null) watch.Validate();
            using (var quiet = new HostQuiet(app)) using (var links = new HostLinkQuiet(app)) using (var bag = new ComBag())
            {
                executionClock.Restart();
                try { calculation = new HostCalculation(app); backup = SummaryWriter.Backup(app, book, Sheet(templateSheet, bag)); ((dynamic)Sheet(templateSheet, bag)).Activate(); }
                catch { EndCalculation(); throw; }
            }
            pending = true; running = true; headersTouched=false; ready = failed = false; step = formulaCount = touched = bulkWrites = 0; view.SetState(true, true, false);
            view.Message("已建立原始已编制区域及目标备份：" + backup + "。逐 Area 执行，不复制外接矩形。"); timer.Start();
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
                    var block = cells.Where(c => c.Row >= sourceArea.Row && c.Row < sourceArea.Row + sourceArea.Rows && c.Column >= sourceArea.Column && c.Column < sourceArea.Column + sourceArea.Columns).ToList();
                    if(block.Any(c=>(!c.Formula&&!c.Literal)||c.Existing!=null))
                    {
                        if(!block.Any(c=>c.Literal)) { target.Formula=SummaryAreaEngine.FormulaMatrix((object)target.Formula,sourceArea,block,view.ProjectFolder,company.Name,targetArea.Row-sourceArea.Row,targetArea.Column-sourceArea.Column);bulkWrites++; }
                        else foreach(var row in block.Where(c=>(!c.Formula&&!c.Literal)||c.Existing!=null).GroupBy(c=>c.Row))
                        {
                            var ordered=row.OrderBy(c=>c.Column).ToList();int first=0;
                            while(first<ordered.Count)
                            {
                                int last=first;while(last+1<ordered.Count&&ordered[last+1].Column==ordered[last].Column+1)last++;
                                var run=new SummaryAreaPlan{Row=row.Key,Column=ordered[first].Column,Rows=1,Columns=ordered[last].Column-ordered[first].Column+1};
                                string address=CellAddress.Name(targetArea.Row+run.Row-sourceArea.Row,targetArea.Column+run.Column-sourceArea.Column)+":"+CellAddress.Name(targetArea.Row+run.Row-sourceArea.Row,targetArea.Column+run.Column-sourceArea.Column+run.Columns-1);
                                dynamic output=bag.Own(sheet.Range[address]);output.Formula=SummaryAreaEngine.FormulaMatrix((object)output.Formula,run,ordered.GetRange(first,last-first+1),view.ProjectFolder,company.Name,targetArea.Row-sourceArea.Row,targetArea.Column-sourceArea.Column);bulkWrites++;first=last+1;
                            }
                        }
                    }
                    block=block.Where(c=>!c.Literal).ToList();
                    if (block.Count > 0)
                    {
                        VerifyFormulaCells(sheet, target, sourceArea, targetArea, block);
                        SummaryAreaEngine.VerifyMatrix((object)target.Formula, (object)target.FormulaR1C1, sourceArea, block, view.ProjectFolder, company.Name,targetArea.Row-sourceArea.Row,targetArea.Column-sourceArea.Column);
                    }
                    formulaCount += block.Count;
                    step++; view.Message("已完成 " + step + "/" + targets.Count + "：" + company.Name + " → " + SummaryAreaEngine.Address(targetArea));
                    if (step == targets.Count)
                    {
                        WriteHeaders(sheet, bag);
                        int errors = 0;
                        foreach (var area in targets)
                        {
                            dynamic range = bag.Own(sheet.Range[SummaryAreaEngine.Address(area)]);
                            try { range.Calculate(); } catch (Exception error) { view.Message("重新计算失败：" + error.Message); }
                            object values = range.Value2;
                            for (int r = 0; r < area.Rows; r++) for (int c = 0; c < area.Columns; c++)
                                if (SummaryAreaEngine.IsError(SummaryArrays.Pick(values, r, area.Rows, c, area.Columns))) errors++;
                        }
                        EndCalculation();
                        view.Message("批量公式提交 " + bulkWrites + " 次；备份及执行耗时 " + executionClock.Elapsed.TotalSeconds.ToString("F2", CultureInfo.InvariantCulture) + " 秒。");
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
            finally
            {
                ticking = false;
                if (!running) { try { EndCalculation(); } catch (Exception error) { view.Message("恢复计算模式失败，请检查 Excel 计算选项：" + error.Message); } executionClock.Stop(); }
                hostBusy(running); if (running) timer.Start();
            }
        }
        private void EndCalculation()
        {
            if (calculation == null) return;
            var scope = calculation; calculation = null;
            using (var links = new HostLinkQuiet(app)) scope.Dispose();
        }
        private static void VerifyFormulaCells(dynamic sheet, dynamic target, SummaryAreaPlan source, SummaryAreaPlan destination, IList<SummaryAreaCell> block)
        {
            // One HasFormula call for dense blocks; for sparse blocks verify
            // horizontal runs, never mistake a literal '=...' for a formula.
            if (Bool((object)target.HasFormula)) return;
            using (var bag = new ComBag())
            foreach (var row in block.GroupBy(c => c.Row))
            {
                var columns = row.Select(c => c.Column).OrderBy(c => c).ToList();
                int start = columns[0], end = start;
                for (int i = 1; i <= columns.Count; i++)
                {
                    if (i < columns.Count && columns[i] == end + 1) { end = columns[i]; continue; }
                    string address = CellAddress.Name(destination.Row + row.Key - source.Row, destination.Column + start - source.Column) + ":" + CellAddress.Name(destination.Row + row.Key - source.Row, destination.Column + end - source.Column);
                    dynamic run = bag.Own(sheet.Range[address]);
                    if (!Bool((object)run.HasFormula)) throw new InvalidOperationException("真实公式写入失败：" + address);
                    if (i < columns.Count) start = end = columns[i];
                }
            }
        }
        private static bool Equivalent(string a, string b) { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }
        private void Restore(int count)
        {
            CheckBook(); using (var quiet = new HostQuiet(app)) using (var links = new HostLinkQuiet(app)) using (var recalc = new HostCalculation(app)) using (var bag = new ComBag())
            {
                dynamic sheet = Sheet(templateSheet, bag), original = Sheet(backup, bag);
                RestoreHeaders(sheet, original, bag);
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
        { if (!pending || running) return; Restore(touched); pending = ready = false; view.SetState(false, false, false); view.Message("已撤销全部受影响目标 Area 和已写入表头，未选中间区域未修改；备份保留，未保存。"); }
        private void Confirm()
        {
            if (!pending || running) return;
            if (failed || step != targets.Count) throw new InvalidOperationException("失败批次只能撤销，不能确认。");
            pending = ready = false; view.SetState(false, false, false); view.Message("已确认并结束本批回滚会话；未保存工作簿。备份仍保留：" + backup);
        }
        private void ReviewReferences(List<string> errors)
        {
            var observations=new Dictionary<string,List<string>>(StringComparer.OrdinalIgnoreCase);var files=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            foreach(var cell in cells.Where(c=>c.Existing!=null))
            foreach(var span in cell.Existing.Spans.Where(x=>x.Kind!=RefKind.Local))
            {
                string key=SummaryExistingFormula.Source(cell.Text,span);if(!observations.ContainsKey(key))observations[key]=new List<string>();
                string item=CellAddress.Name(cell.Row,cell.Column)+" | "+cell.Text;if(!observations[key].Contains(item))observations[key].Add(item);files[key]=SummaryExistingFormula.File(cell.Text,span);
            }
            var review=new List<SummaryReferenceSource>();
            foreach(var entry in observations)
            {
                bool matched=files[entry.Key].Length>0&&companies.Any(c=>string.Equals(c.Name+".xlsx",files[entry.Key],StringComparison.OrdinalIgnoreCase));
                review.Add(new SummaryReferenceSource{Key=entry.Key,File=files[entry.Key],Locations=string.Join(", ",entry.Value.Select(x=>x.Split('|')[0].Trim()).Distinct().ToArray()),Signature=string.Join("\n",entry.Value.ToArray()),Reason=matched?"匹配 Lead 公司文件" : files[entry.Key].Length==0?"本工作簿跨表来源，只能保留" : "未匹配 Lead，可能为公共或无关来源",Decision=matched&&observations.Count==1?"参与公司替换":"待确认"});
                if(observations.Count>1)review.Last().Reason+="；存在多个来源，须逐项确认";
                view.Message("来源预检："+entry.Key+"；"+review.Last().Reason);foreach(string item in entry.Value)view.Message("引用位置及公式："+item);
            }
            var decisions=view.ReviewSources(review);
            foreach(var source in review)
            {
                string decision=decisions[source.Key];
                if(decision=="待确认"||string.IsNullOrEmpty(decision)){errors.Add("引用来源待确认："+source.Key+"；请在来源清单指定处理方式，再重新预检。");continue;}
                if(decision=="参与公司替换"&&source.File.Length==0){errors.Add("本工作簿跨表来源只能选择保留原引用："+source.Key);continue;}
                foreach(var cell in cells.Where(c=>c.Existing!=null))cell.Existing.Decisions[source.Key]=decision=="参与公司替换";
                view.Message(source.Key+" → "+decision);
            }
        }
        private void PrepareHeaders(ComBag bag, IList<SummaryCompany> picked)
        {
            headerTargets.Clear();headerValues.Clear();view.HeaderPreview.Text="";
            if(!view.HeadersEnabled.Checked)return;
            if(companyHeader==null||yearSource==null||yearTarget==null)throw new InvalidOperationException("启用表头填充后，三个新增选区必须完整。");
            if(companyHeader.Sheet!=templateSheet||yearTarget.Sheet!=templateSheet)throw new InvalidOperationException("表头目标必须与已编制区域、粘贴起点同表。");
            int width=areas.Max(a=>a.Column+a.Columns)-areas.Min(a=>a.Column);
            if(yearSource.Rows!=1||yearSource.Columns!=width)throw new InvalidOperationException("年度来源有 "+yearSource.Columns+" 格，已编制区域横向跨度为 "+width+" 列，必须相等。");
            dynamic years=Sheet(yearSource.Sheet,bag), sheet=Sheet(templateSheet,bag);
            var values=new List<string>();
            for(int c=0;c<width;c++)using(var one=new ComBag()) { dynamic cell=one.Own(years.Range[CellAddress.Name(yearSource.Row,yearSource.Column+c)]);RejectSpecial(cell,"年度来源",one); string text=Convert.ToString(cell.Text);if(string.IsNullOrWhiteSpace(text)||text.Trim('#').Length==0||SummaryAreaEngine.IsError((object)cell.Value2))throw new InvalidOperationException("年度显示值为空、错误或列宽不足，请检查来源。"); values.Add(text); }
            view.Message("表头预检：公司 "+picked.Count+" 家；已编制区域跨度 "+width+"；年度 "+yearSource.Sheet+"!"+SummaryAreaEngine.Address(yearSource)+"："+string.Join("、",values.ToArray()));
            for(int i=0;i<picked.Count;i++)
            {
                long nameColumn=(long)companyHeader.Column+i*width, yearColumn=(long)yearTarget.Column+i*width;
                if(nameColumn>16384||yearColumn+width-1>16384)throw new InvalidOperationException("表头映射超出 Excel 列边界。");
                using(var one=new ComBag())
                {
                    dynamic cell=one.Own(sheet.Range[CellAddress.Name(companyHeader.Row,(int)nameColumn)]);
                    var name=new SummaryAreaPlan{Sheet=templateSheet,Row=companyHeader.Row,Column=(int)nameColumn,Rows=1,Columns=1};
                    if(Bool((object)cell.MergeCells)) {dynamic merge=one.Own(cell.MergeArea);name=Coordinates(merge,one);name.Sheet=templateSheet;}
                    dynamic nameRange=one.Own(sheet.Range[SummaryAreaEngine.Address(name)]); RejectArraySpill(nameRange,"公司名称目标",one);
                    int groupStart=paste.Column+i*width;
                    if(name.Column<groupStart||(long)name.Column+name.Columns>groupStart+width)throw new InvalidOperationException("公司名称合并区域跨越公司分组或授权边界："+SummaryAreaEngine.Address(name));
                    AddHeader(name,picked[i].Code);
                    for(int c=0;c<width;c++)
                    {
                        var year=new SummaryAreaPlan{Sheet=templateSheet,Row=yearTarget.Row,Column=(int)yearColumn+c,Rows=1,Columns=1};dynamic y=one.Own(sheet.Range[CellAddress.Name(year.Row,year.Column)]);RejectSpecial(y,"年度目标",one);AddHeader(year,values[c]);
                    }
                    view.Message((i+1)+". "+picked[i].Code+" → 名称 "+SummaryAreaEngine.Address(name)+"；年度 "+CellAddress.Name(yearTarget.Row,(int)yearColumn)+":"+CellAddress.Name(yearTarget.Row,(int)yearColumn+width-1));
                }
            }
            view.HeaderPreview.Text=picked.Count+" 家公司，名称 "+picked.Count+" 处，年度 "+(picked.Count*width)+" 格："+string.Join("、",values.ToArray());
            foreach(var part in headerTargets)
            {
                dynamic range=bag.Own(sheet.Range[SummaryAreaEngine.Address(part)]);view.Message("表头授权覆盖 "+SummaryAreaEngine.Address(part)+"；合并="+Convert.ToString(range.MergeCells)+"；原内容="+Convert.ToString(((dynamic)bag.Own(sheet.Range[CellAddress.Name(part.Row,part.Column)])).Formula));
                dynamic first=bag.Own(sheet.Range[CellAddress.Name(part.Row,part.Column)]);if(Bool((object)first.HasFormula))originalTargetFormulas.Add(new SummaryAreaCell{Row=part.Row,Column=part.Column,Text=Convert.ToString(first.Formula)});
            }
        }
        private void AddHeader(SummaryAreaPlan part,string value)
        {
            if(areas.Any(a=>SummaryAreaEngine.Intersects(a,part))||targets.Any(a=>SummaryAreaEngine.Intersects(a,part))||headerTargets.Any(a=>SummaryAreaEngine.Intersects(a,part))||(yearSource.Sheet==part.Sheet&&SummaryAreaEngine.Intersects(yearSource,part))||(leadSheet==part.Sheet&&LeadSelections().Any(a=>SummaryAreaEngine.Intersects(a,part))))throw new InvalidOperationException("表头目标与已编制区域、复制目标、其他表头或来源选区冲突："+SummaryAreaEngine.Address(part));
            headerTargets.Add(part);headerValues.Add(value);
        }
        private void WriteHeaders(dynamic sheet,ComBag bag)
        {
            headersTouched=headerTargets.Count>0;
            for(int i=0;i<headerTargets.Count;i++)
            {
                var part=headerTargets[i];dynamic cell=bag.Own(sheet.Range[CellAddress.Name(part.Row,part.Column)]);object format=cell.NumberFormat;
                try{cell.NumberFormat="@";cell.Value2=headerValues[i];}finally{cell.NumberFormat=format;}
                if(Bool((object)cell.HasFormula)||Convert.ToString(cell.Value2)!=headerValues[i])throw new InvalidOperationException("表头文字读回校验失败："+SummaryAreaEngine.Address(part));
            }
            if(headerTargets.Count>0)view.Message("公司名称与年度表头写入、读回校验成功 "+headerTargets.Count+" 项。");
        }
        private void RestoreHeaders(dynamic sheet,dynamic original,ComBag bag)
        {
            if(!headersTouched)return;
            foreach(var part in headerTargets)
            {
                string address=SummaryAreaEngine.Address(part);dynamic source=bag.Own(original.Range[address]),destination=bag.Own(sheet.Range[address]);source.Copy(destination);
                foreach(var prior in originalTargetFormulas.Where(p=>p.Row>=part.Row&&p.Row<part.Row+part.Rows&&p.Column>=part.Column&&p.Column<part.Column+part.Columns)){dynamic cell=bag.Own(sheet.Range[CellAddress.Name(prior.Row,prior.Column)]);cell.Formula=prior.Text;if(Convert.ToString(cell.Formula)!=prior.Text)throw new InvalidOperationException("表头原公式恢复失败。");}
            }
        }
        public void Dispose()
        {
            timer.Stop(); timer.Dispose();
            try { EndCalculation(); } catch (Exception error) { UiDiagnostics.Record("summary-calculation-restore", error); }
            finally { hostBusy(false); if (presetBinding != null) presetBinding.Dispose(); if (watch != null) watch.Dispose(); references.Dispose(); }
        }
    }
}
