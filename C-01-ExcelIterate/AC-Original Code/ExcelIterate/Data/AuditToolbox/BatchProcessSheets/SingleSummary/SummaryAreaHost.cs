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
        private readonly List<SummaryMappedWork> mappedWork=new List<SummaryMappedWork>();
        private readonly List<SummaryAreaPlan> fixedReadRanges=new List<SummaryAreaPlan>();
        private readonly Dictionary<int,Dictionary<int,int>> fixedColumnMaps=new Dictionary<int,Dictionary<int,int>>();
        private int fixedNameRow,fixedYearRow;
        private readonly Dictionary<string,string> fixedYearCache=new Dictionary<string,string>();

        private SummaryAreaPlan companyHeader, yearSource, yearTarget;
        private readonly List<SummaryAreaPlan> headerTargets = new List<SummaryAreaPlan>();
        private readonly List<string> headerValues = new List<string>();
        private bool headersTouched;
        private readonly Dictionary<string,string> headerExpressions=new Dictionary<string,string>(),headerFormats=new Dictionary<string,string>();
        private readonly List<SummaryHeaderArray> oldHeaderArrays=new List<SummaryHeaderArray>(),yearHeaderArrays=new List<SummaryHeaderArray>();


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
            this.approveOverwrite = approveOverwrite ?? delegate { return UiDialog.Show(form, view.MatchHeaders.Checked?"将按预检的公司＋年度列名覆盖目标数据（含选区空白格）。现有表头、来源及未选区域不改动。确认执行？":"将覆盖预检日志中的目标 Area 和启用的表头。确认执行？", "确认覆盖", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.OK; };
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
            preset.Values["matchHeaders"]=view.MatchHeaders.Checked.ToString();
            preset.Values["headersEnabled"] = view.HeadersEnabled.Checked.ToString();preset.Values["headersFormula"]=view.FormulaHeaders.Checked.ToString();preset.Values["headersRepeat"]=view.RepeatNames.Checked.ToString();
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
                    foreach (var part in settings.Template.Concat(settings.Lead).Concat(settings.Paste==null?new SummaryAreaPlan[0]:new[] { settings.Paste }))
                    { dynamic sheet = bag.Own(sheets.Item(part.Sheet)); dynamic range = bag.Own(sheet.Range[SummaryAreaEngine.Address(part)]); if(!settings.MatchHeaders||settings.Lead.Contains(part))RejectSpecial(range, part.Address, bag); }
                }
                using (var bag = new ComBag())
                {
                    dynamic sheets = bag.Own(((dynamic)selectedBook).Worksheets);
                    foreach(var part in new[]{settings.CompanyHeader,settings.YearSource,settings.YearTarget}.Where(p=>p!=null&&!settings.MatchHeaders))
                    { dynamic sheet = bag.Own(sheets.Item(part.Sheet)); dynamic range = bag.Own(sheet.Range[SummaryAreaEngine.Address(part)]); if(!settings.FormulaHeaders){if(part==settings.CompanyHeader)RejectArraySpill(range,"公司名称",bag);else RejectSpecial(range,"表头选区",bag);}
                        else if(part!=settings.CompanyHeader){object merged=range.MergeCells;if(merged==null||merged==DBNull.Value||Bool(merged))throw new InvalidOperationException("年度选区不能包含合并格。");} }
                }
                var oldBinding = presetBinding;
                book = selectedBook; bookName = Convert.ToString(((dynamic)book).FullName); presetBinding = candidate;
                templateSheet = settings.Template[0].Sheet; pasteSheet = settings.Paste==null?templateSheet:settings.Paste.Sheet; leadSheet = settings.Lead[0].Sheet;
                areas.Clear(); areas.AddRange(settings.Template); paste = settings.Paste; leadAreas.Clear(); leadAreas.AddRange(settings.Lead); lead = leadAreas[0];
                loading = true;
                try
                {
                    view.MatchHeaders.Checked=settings.MatchHeaders;companyHeader=settings.CompanyHeader; yearSource=settings.YearSource; yearTarget=settings.YearTarget; view.HeadersEnabled.Checked=settings.HeadersEnabled;view.FormulaHeaders.Checked=settings.FormulaHeaders;view.RepeatNames.Checked=settings.RepeatNames;view.ReplaceHeaderArrays.Checked=false;
                    var headerNames=new[]{"companyHeader","yearSource","yearTarget"}; var headerParts=new[]{companyHeader,yearSource,yearTarget};
                    for(int i=0;i<headerNames.Length;i++) view.Address(headerNames[i],headerParts[i]==null ? "未选择" : headerParts[i].Sheet+"!"+SummaryAreaEngine.Address(headerParts[i]));
                    view.HeaderPreview.Text="请重新预检以生成映射预览。";
                    view.Folder.Text = settings.Folder; view.CheckFiles.Checked = settings.CheckFiles; view.Companies.Rows.Clear(); view.ReferenceSources.Rows.Clear();
                    view.Address("template", templateSheet + "!" + string.Join("; ", areas.Select(a => a.Address).ToArray()));
                    view.Address("paste", paste==null?"按列名自动定位":pasteSheet + "!" + paste.Address); view.Address("lead", leadSheet + "!" + string.Join("; ", leadAreas.Select(a => a.Address).ToArray()));
                }
                finally { loading = false; }
                ready = false; companies.Clear(); cells.Clear(); targets.Clear();mappedWork.Clear();fixedReadRanges.Clear();fixedColumnMaps.Clear(); committed = true;
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
                        if (native) { if (kind == "lead" || kind == "yearSource") ActivateLeadSheet(); form.Hide(); chosen = DataMigrationHost.PickNativeRange(app, kind == "template" ? "选择已编制区域，可按 Ctrl 多选；若原生框不支持，请取消后使用读取当前选区。" : kind == "lead" ? "选择编码和全称两列的数据行，可 Ctrl 多选分段或不相邻两列；同一行需配对，不含表头。" : kind == "yearSource" ? "选择单列或单行连续年度来源；单列转置成一行，数量须等于已编制区域横向跨度。" : kind == "companyHeader" ? "选择第一家公司名称目标单元格，支持既有合并区域，不创建或拆分合并。" : kind == "yearTarget" ? "选择第一家公司年度写入起点，一个单元格。" : "选择粘贴起点，一个单元格。", "单体汇总批量链接", Type.Missing); }
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
                        if (kind == "yearSource" && values[0].Rows!=1 && values[0].Columns!=1)throw new InvalidOperationException("年度来源只接受单列或单行连续区域。");
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
        // Range-level Null means a mixed result, not an affirmative spill.
        // Recheck physical cells; merged cells share the actual write anchor.
        private void RejectArraySpill(dynamic range,string address,ComBag bag)
        {
            CheckFormulaFlag(range,address,false,bag);
            CheckFormulaFlag(range,address,true,bag);
        }
        internal static int FormulaFlag(object value,string address,string property)
        {
            if(value==null||value==DBNull.Value)return -1;
            if(value is bool)return (bool)value?1:0;
            if(value is short||value is int)
            {
                int number=Convert.ToInt32(value);if(number==0)return 0;if(number==1||number==-1)return 1;
            }
            throw new InvalidOperationException(address+" 无法确认 "+property+" 状态；返回类型="+value.GetType().Name+"，值="+Convert.ToString(value,CultureInfo.InvariantCulture)+"。未执行写入。");
        }
        private void CheckFormulaFlag(dynamic range,string address,bool spill,ComBag bag)
        {
            string property=spill?"HasSpill":"HasArray",description=spill?"动态溢出":"数组";
            object aggregate;
            try { aggregate=spill?(object)range.HasSpill:(object)range.HasArray; }
            catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { if(!spill)throw; return; }
            catch(System.Runtime.InteropServices.COMException error)
            { if(spill&&(error.HResult==unchecked((int)0x80020003)||error.HResult==unchecked((int)0x80020006)))return;throw; }
            // Only an unsupported aggregate property identifies an old host.
            // A per-cell failure after an available aggregate must remain blocking.
            int state=FormulaFlag(aggregate,address,property);
            if(state==1)throw new InvalidOperationException(address+" 含"+description+"公式（"+property+"=True）。");
            if(state==0)return;
            var bounds=Coordinates(range,bag);dynamic cells=bag.Own(range.Cells);
            for(int r=1;r<=bounds.Rows;r++)for(int c=1;c<=bounds.Columns;c++)
            using(var one=new ComBag())
            {
                dynamic cell=one.Own(cells.Item(r,c)),anchor=cell;
                if(Bool((object)cell.MergeCells)) { dynamic merge=one.Own(cell.MergeArea),mergedCells=one.Own(merge.Cells);anchor=one.Own(mergedCells.Item(1,1)); }
                string location=address+" / "+Convert.ToString(anchor.Address);
                int single=FormulaFlag(spill?(object)anchor.HasSpill:(object)anchor.HasArray,location,property);
                if(single==1)throw new InvalidOperationException(location+" 含"+description+"公式（"+property+"=True）。");
                if(single<0)throw new InvalidOperationException(location+" 单格 "+property+" 仍返回混合或空状态，无法安全确认；未执行写入。请提供该位置及 Excel/WPS 版本以便诊断。");
            }
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
                    companies.Add(company); view.Companies.Rows.Add(company.Selected, company.Row, code, name,company.Selected?"待预检":"未勾选");
                }
            }
            finally { loading = false; }
        }
        private sealed class FixedPeriod { internal string Key;internal object Value;internal int Row; }
        private sealed class FixedScanLine { internal int Row;internal object Values; }
        private sealed class FixedHeaderPair
        {
            internal int NameRow,YearRow,Company;
            internal long Score;
            internal Dictionary<int,string> Years;
        }
        private static string FixedText(object value){return Convert.ToString(value,CultureInfo.InvariantCulture).Trim();}
        private Dictionary<string,List<int>> CompanyLookup()
        {
            var lookup=new Dictionary<string,List<int>>(StringComparer.OrdinalIgnoreCase);
            foreach(var company in companies)foreach(string raw in new[]{company.Code,company.Name})
            {
                string text=FixedText(raw);if(text.Length==0)continue;if(!lookup.ContainsKey(text))lookup[text]=new List<int>();if(!lookup[text].Contains(company.Row))lookup[text].Add(company.Row);
            }
            return lookup;
        }
        private static int[] FixedNames(object values,int columns,int years,IDictionary<string,List<int>> lookup)
        {
            var result=new int[columns];int current=0,start=0;
            for(int c=0;c<columns;c++)
            {
                object raw=SummaryArrays.Pick(values,0,1,c,columns);string text=FixedText(raw);
                if(text.Length>0)
                {
                    List<int> rows;int found=raw is string&&lookup.TryGetValue(text,out rows)&&rows.Count==1?rows[0]:0;
                    if(found==0){current=0;start=c;}else {if(found!=current||c-start>=years)start=c;current=found;}
                }
                else if(c-start>=years)current=0;
                result[c]=current;
            }
            return result;
        }
        private List<FixedPeriod> ReadFixedPeriods(ComBag bag)
        {
            dynamic sheet=Sheet(leadSheet,bag),used=bag.Own(sheet.UsedRange);int last=Math.Min(10000,Convert.ToInt32(used.Row)+Convert.ToInt32(((dynamic)bag.Own(used.Rows)).Count)-1);
            dynamic range=bag.Own(sheet.Range["N1:N"+Math.Max(1,last)]);object values=range.Value2;int heading=0;
            var titles=new HashSet<string>(new[]{"期间","年度","年份","时间","Period","Year"},StringComparer.OrdinalIgnoreCase);
            for(int r=1;r<=last;r++)if(titles.Contains(FixedText(SummaryArrays.Pick(values,r-1,last,0,1)))){heading=r;break;}
            if(heading==0)throw new InvalidOperationException("公司来源表 N 列未找到期间/年度表头；请检查现有年度配置，未修改单元格。");
            var result=new List<FixedPeriod>();
            for(int r=heading+1;r<=last&&result.Count<=100;r++)
            {
                object value=SummaryArrays.Pick(values,r-1,last,0,1);if(!SummaryArrays.HasContent(value))break;
                dynamic cell=bag.Own(sheet.Range["N"+r]);string key=FixedText(cell.Text);
                if(SummaryAreaEngine.IsError(value)||key.Length==0||key.Trim('#').Length==0)throw new InvalidOperationException("年度来源显示错误或列宽不足："+leadSheet+"!N"+r);
                result.Add(new FixedPeriod{Key=key,Value=value,Row=r});
            }
            if(result.Count==0||result.Count>100||result.Select(y=>y.Key).Distinct(StringComparer.Ordinal).Count()!=result.Count)throw new InvalidOperationException("N 列期间为空、过长或重复，无法唯一匹配年度列名。");
            fixedReadRanges.Add(new SummaryAreaPlan{Sheet=leadSheet,Row=heading,Column=14,Rows=result.Count+1,Columns=1});return result;
        }
        private string FixedYear(object value,dynamic sheet,int row,int column,IList<FixedPeriod> years,ComBag bag)
        {
            if(SummaryAreaEngine.IsError(value))return "";
            if(value is string){string raw=FixedText(value);return years.Any(y=>y.Key==raw)?raw:"";}
            if(!SummaryArrays.HasContent(value))return "";
            string address=row+":"+column,display;if(!fixedYearCache.TryGetValue(address,out display))
            {using(var one=new ComBag()){dynamic cell=one.Own(sheet.Range[CellAddress.Name(row,column)]);display=FixedText(cell.Text);}fixedYearCache[address]=display;}
            return years.Any(y=>y.Key==display)?display:"";
        }
        private void PrepareColumnMatches(dynamic sheet,ComBag bag)
        {
            fixedYearCache.Clear();SummaryAreaEngine.ValidateAreas(areas);if(areas.Any(a=>a.Sheet!=templateSheet))throw new InvalidOperationException("已编制区域不能跨工作表。");
            var years=ReadFixedPeriods(bag);var lookup=CompanyLookup();int first=areas.Min(a=>a.Column),last=areas.Max(a=>a.Column+a.Columns-1),top=areas.Min(a=>a.Row);
            if(last-first+1>years.Count)throw new InvalidOperationException("请选择一家公司已编制的数据列；来源跨度 "+(last-first+1)+"，期间数 "+years.Count+"。");
            var selectedColumns=areas.SelectMany(a=>Enumerable.Range(a.Column,a.Columns)).Distinct().ToList();
            dynamic used=bag.Own(sheet.UsedRange);int lastRow=Convert.ToInt32(used.Row)+Convert.ToInt32(((dynamic)bag.Own(used.Rows)).Count)-1,lastColumn=Convert.ToInt32(used.Column)+Convert.ToInt32(((dynamic)bag.Own(used.Columns)).Count)-1;
            int scanFirst=Math.Max(1,first-years.Count+1),scanLast=Math.Min(16384,last+years.Count-1),columns=scanLast-scanFirst+1;
            var lines=new Dictionary<int,FixedScanLine>();var bands=new[]{new[]{1,Math.Min(lastRow,128)},new[]{Math.Max(1,top-64),Math.Min(lastRow,top+64)}};
            foreach(var band in bands)
            {
                if(band[1]<band[0])continue;int rows=band[1]-band[0]+1;dynamic range=bag.Own(sheet.Range[CellAddress.Name(band[0],scanFirst)+":"+CellAddress.Name(band[1],scanLast)]);object data=range.Value2;
                for(int r=0;r<rows;r++){var values=new object[1,columns];for(int c=0;c<columns;c++)values[0,c]=SummaryArrays.Pick(data,r,rows,c,columns);lines[band[0]+r]=new FixedScanLine{Row=band[0]+r,Values=values};}
            }
            var candidates=new List<FixedHeaderPair>();
            foreach(var name in lines.Values)
            {
                int[] ids=FixedNames(name.Values,columns,years.Count,lookup);int owner=ids[first-scanFirst];if(owner==0||selectedColumns.Any(c=>ids[c-scanFirst]!=owner))continue;
                int left=first-scanFirst,right=last-scanFirst;while(left>0&&ids[left-1]==owner)left--;while(right+1<columns&&ids[right+1]==owner)right++;
                foreach(var year in lines.Values.Where(y=>y.Row>name.Row))
                {
                    var sourceYears=new Dictionary<int,string>();for(int c=left;c<=right;c++){string key=FixedYear(SummaryArrays.Pick(year.Values,0,1,c,columns),sheet,year.Row,scanFirst+c,years,bag);if(key.Length>0)sourceYears[scanFirst+c]=key;}
                    if(sourceYears.Count!=years.Count||sourceYears.Values.Distinct(StringComparer.Ordinal).Count()!=years.Count||selectedColumns.Any(c=>!sourceYears.ContainsKey(c)))continue;
                    candidates.Add(new FixedHeaderPair{NameRow=name.Row,YearRow=year.Row,Company=owner,Years=sourceYears,Score=Math.Abs((long)year.Row-top)*2048+year.Row-name.Row});
                }
            }
            if(candidates.Count==0)throw new InvalidOperationException("未识别有效的公司/年度列名；请检查现有 LET 结果、N 列期间和所选数据列。表头若有 #REF! 须先修复，本模式不重建表头。");
            var pair=candidates.OrderBy(c=>c.Score).First();if(candidates.Count(c=>c.Score==pair.Score)>1)throw new InvalidOperationException("识别到多个同等候选表头，无法安全定位；请把实际表头一同纳入已编制区域后重新预检。");
            fixedNameRow=pair.NameRow;fixedYearRow=pair.YearRow;
            dynamic names=bag.Own(sheet.Range["A"+pair.NameRow+":"+CellAddress.Name(pair.NameRow,lastColumn)]),periods=bag.Own(sheet.Range["A"+pair.YearRow+":"+CellAddress.Name(pair.YearRow,lastColumn)]);
            object nameValues=names.Value2,periodValues=periods.Value2;int[] companyIds=FixedNames(nameValues,lastColumn,years.Count,lookup);var headers=new List<SummaryColumnHeader>();
            for(int c=1;c<=lastColumn;c++)
            {
                if(companyIds[c-1]==0)continue;string key=FixedYear(SummaryArrays.Pick(periodValues,0,1,c-1,lastColumn),sheet,pair.YearRow,c,years,bag);if(key.Length>0)headers.Add(new SummaryColumnHeader{CompanyRow=companyIds[c-1],Column=c,Year=key});
            }
            foreach(var company in companies)
            {
                if(company.Row==pair.Company){company.Selected=false;view.Message("已有公司，整组跳过："+company.Code+"；来源列 "+CellAddress.ColumnLetters(first)+":"+CellAddress.ColumnLetters(last));}
                foreach(DataGridViewRow row in view.Companies.Rows)if(Convert.ToInt32(row.Cells["row"].Value)==company.Row)row.Cells["status"].Value=company.Row==pair.Company?"已编制，跳过":company.Selected?"将生成":"未勾选";
            }
            var maps=SummaryColumnMapping.Build(companies.Where(c=>c.Selected).Select(c=>c.Row),pair.Company,pair.Years,headers);foreach(var entry in maps)fixedColumnMaps[entry.Key]=entry.Value;
            mappedWork.AddRange(SummaryColumnMapping.Work(areas,maps,new HashSet<int>{pair.NameRow,pair.YearRow}));targets=mappedWork.Select(w=>w.Target).ToList();
            fixedReadRanges.Add(new SummaryAreaPlan{Sheet=templateSheet,Row=pair.NameRow,Column=1,Rows=1,Columns=lastColumn});fixedReadRanges.Add(new SummaryAreaPlan{Sheet=templateSheet,Row=pair.YearRow,Column=1,Rows=1,Columns=lastColumn});
            view.Message("复用现有表头：公司行 "+pair.NameRow+"，年度行 "+pair.YearRow+"；数据保持原行号；名称、年度及其公式均不写入。");
            foreach(var entry in maps){var company=companies.Single(c=>c.Row==entry.Key);foreach(var col in entry.Value)view.Message(company.Code+" / "+pair.Years[col.Key]+"："+CellAddress.ColumnLetters(col.Key)+" → "+CellAddress.ColumnLetters(col.Value));}
        }
        private SummaryCompany CompanyForStep(int index,IList<SummaryCompany> picked)
        {return mappedWork.Count>0?companies.Single(c=>c.Row==mappedWork[index].CompanyRow):picked[index/areas.Count];}
        private IDictionary<int,int> ColumnMap(int company)
        {Dictionary<int,int> map;return view.MatchHeaders.Checked&&fixedColumnMaps.TryGetValue(company,out map)?map:null;}
        private void Preview()
        {
            if (running || pending) return; ready = false; CheckBook();
            if (areas.Count == 0 || (!view.MatchHeaders.Checked&&paste == null) || lead == null) throw new InvalidOperationException("请完成三组选区。");
            if (!view.MatchHeaders.Checked&&templateSheet != pasteSheet) throw new InvalidOperationException("已编制区域和粘贴起点必须属于同一工作表。");
            if (!Directory.Exists(view.ProjectFolder)) throw new InvalidOperationException("项目文件夹不存在。");
            using (var bag = new ComBag())
            {
                if (Bool(((dynamic)book).ReadOnly) || Bool(((dynamic)book).ProtectStructure)) throw new InvalidOperationException("工作簿只读或结构受保护，不能执行及备份。");
                dynamic sheet = Sheet(templateSheet, bag);
                if (Bool(sheet.ProtectContents)) throw new InvalidOperationException("目标工作表受保护。");
                ReadCompanies(bag);mappedWork.Clear();fixedReadRanges.Clear();fixedColumnMaps.Clear();
                if(view.MatchHeaders.Checked)PrepareColumnMatches(sheet,bag);
                var picked = companies.Where(c => c.Selected).ToList();
                if(!view.MatchHeaders.Checked)targets=SummaryAreaEngine.Targets(areas,paste.Row,paste.Column,picked.Count);
                var scanAreas=view.MatchHeaders.Checked?SummaryColumnMapping.DataAreas(areas,new HashSet<int>{fixedNameRow,fixedYearRow}):areas;
                if (leadSheet == templateSheet && targets.Any(t => LeadSelections().Any(part => SummaryAreaEngine.Intersects(t, part)))) throw new InvalidOperationException("目标不能覆盖 Lead 公司字段。");
                cells.Clear(); var errors = new List<string>();
                foreach (var area in scanAreas)
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
                                if (formula) { existing=SummaryExistingFormula.Read(text); if(!view.MatchHeaders.Checked&&!existing.Spans.Any(x=>x.Kind!=RefKind.Local))existing=null; }
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
                    var source = mappedWork.Count>0?mappedWork[i].Source:areas[i % areas.Count]; var target = targets[i];
                    foreach (var cell in cells.Where(c => c.Formula && c.Row >= source.Row && c.Row < source.Row + source.Rows && c.Column >= source.Column && c.Column < source.Column + source.Columns))
                        try { if(cell.Existing==null)SummaryAreaEngine.ValidateShift(cell.Text, target.Row - source.Row, target.Column - source.Column); else cell.Existing.Render(view.ProjectFolder,CompanyForStep(i,picked).Name,target.Row-source.Row,target.Column-source.Column,ColumnMap(CompanyForStep(i,picked).Row)); }
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
        { return view.MatchHeaders.Checked+"|"+view.FormulaHeaders.Checked+"|"+view.RepeatNames.Checked+"|"+view.ReplaceHeaderArrays.Checked+"|"+string.Join(";",view.ReferenceSources.Rows.Cast<DataGridViewRow>().Select(r=>Convert.ToString(r.Cells["source"].Value)+":"+Convert.ToString(r.Cells["decision"].Value)).ToArray()) + "|" + view.HeadersEnabled.Checked + "|" + string.Join(";",new[]{companyHeader,yearSource,yearTarget}.Select(p=>p==null ? "" : p.Sheet+"!"+SummaryAreaSettings.Encode(new[]{p})).ToArray()) + "|" + view.Folder.Text + "|" + view.CheckFiles.Checked + "|" + string.Join(",", view.Companies.Rows.Cast<DataGridViewRow>().Select(r => Convert.ToString(r.Cells["row"].Value) + ":" + Convert.ToString(r.Cells["selected"].Value)).ToArray()); }
        private string Fingerprint()
        {
            CheckBook(); var clock = System.Diagnostics.Stopwatch.StartNew(); var text = new StringBuilder(bookName); int xmlAreas = 0, fallbackAreas = 0;
            using (var bag = new ComBag())
            {
                var leadParts = LeadSelections();
                if(view.MatchHeaders.Checked){dynamic table=Sheet(templateSheet,bag),used=bag.Own(table.UsedRange);Append(text,used.Row);Append(text,used.Column);Append(text,((dynamic)bag.Own(used.Rows)).Count);Append(text,((dynamic)bag.Own(used.Columns)).Count);}
                foreach (var part in areas.Concat(targets).Concat(leadParts).Concat(headerTargets).Concat(fixedReadRanges).Concat(oldHeaderArrays.Select(a=>a.Area)).Concat(view.HeadersEnabled.Checked && yearSource!=null ? new[]{yearSource} : new SummaryAreaPlan[0]))
                {
                    dynamic sheet = Sheet(leadParts.Contains(part) ? leadSheet : part.Sheet ?? templateSheet, bag);
                    if (Bool(sheet.ProtectContents)) throw new InvalidOperationException("工作表保护状态变化。");
                    using (var snapshot = new ComBag())
                    {
                        dynamic range = snapshot.Own(sheet.Range[SummaryAreaEngine.Address(part)]); string xml = null;
                        // Company names may themselves be formulas. Their evaluated
                        // values determine filenames and must invalidate the plan.
                        if (leadParts.Contains(part) || fixedReadRanges.Contains(part) || (view.HeadersEnabled.Checked && part==yearSource))
                        {
                            object values = range.Value2;
                            for (int r = 0; r < part.Rows; r++) for (int c = 0; c < part.Columns; c++) Append(text, SummaryArrays.Pick(values, r, part.Rows, c, part.Columns));
                        }
                        if((view.HeadersEnabled.Checked && part==yearSource)||fixedReadRanges.Contains(part)) for(int c=0;c<Math.Max(part.Rows,part.Columns);c++) { dynamic displayed=snapshot.Own(sheet.Range[CellAddress.Name(part.Row+(part.Rows>1?c:0),part.Column+(part.Rows==1?c:0))]); Append(text,displayed.Text); }
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
                    var sourceArea=mappedWork.Count>0?mappedWork[step].Source:areas[step%areas.Count];var targetArea=targets[step];var company=CompanyForStep(step,picked);
                    dynamic source = bag.Own(original.Range[SummaryAreaEngine.Address(sourceArea)]), target = bag.Own(sheet.Range[SummaryAreaEngine.Address(targetArea)]);
                    touched = step + 1; source.Copy(target);
                    var block = cells.Where(c => c.Row >= sourceArea.Row && c.Row < sourceArea.Row + sourceArea.Rows && c.Column >= sourceArea.Column && c.Column < sourceArea.Column + sourceArea.Columns).ToList();
                    if(block.Any(c=>(!c.Formula&&!c.Literal)||c.Existing!=null))
                    {
                        if(!block.Any(c=>c.Literal)) { target.Formula=SummaryAreaEngine.FormulaMatrix((object)target.Formula,sourceArea,block,view.ProjectFolder,company.Name,targetArea.Row-sourceArea.Row,targetArea.Column-sourceArea.Column,ColumnMap(company.Row));bulkWrites++; }
                        else foreach(var row in block.Where(c=>(!c.Formula&&!c.Literal)||c.Existing!=null).GroupBy(c=>c.Row))
                        {
                            var ordered=row.OrderBy(c=>c.Column).ToList();int first=0;
                            while(first<ordered.Count)
                            {
                                int last=first;while(last+1<ordered.Count&&ordered[last+1].Column==ordered[last].Column+1)last++;
                                var run=new SummaryAreaPlan{Row=row.Key,Column=ordered[first].Column,Rows=1,Columns=ordered[last].Column-ordered[first].Column+1};
                                string address=CellAddress.Name(targetArea.Row+run.Row-sourceArea.Row,targetArea.Column+run.Column-sourceArea.Column)+":"+CellAddress.Name(targetArea.Row+run.Row-sourceArea.Row,targetArea.Column+run.Column-sourceArea.Column+run.Columns-1);
                                dynamic output=bag.Own(sheet.Range[address]);output.Formula=SummaryAreaEngine.FormulaMatrix((object)output.Formula,run,ordered.GetRange(first,last-first+1),view.ProjectFolder,company.Name,targetArea.Row-sourceArea.Row,targetArea.Column-sourceArea.Column,ColumnMap(company.Row));bulkWrites++;first=last+1;
                            }
                        }
                    }
                    block=block.Where(c=>!c.Literal).ToList();
                    if (block.Count > 0)
                    {
                        VerifyFormulaCells(sheet, target, sourceArea, targetArea, block);
                        SummaryAreaEngine.VerifyMatrix((object)target.Formula, (object)target.FormulaR1C1, sourceArea, block, view.ProjectFolder, company.Name,targetArea.Row-sourceArea.Row,targetArea.Column-sourceArea.Column,ColumnMap(company.Row));
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
        private static string HeaderKey(SummaryAreaPlan part){return part.Row+":"+part.Column;}
        private void PrepareHeaders(ComBag bag, IList<SummaryCompany> picked)
        {
            headerTargets.Clear();headerValues.Clear();headerExpressions.Clear();headerFormats.Clear();oldHeaderArrays.Clear();yearHeaderArrays.Clear();view.HeaderPreview.Text="";
            if(view.MatchHeaders.Checked){view.HeaderPreview.Text="复用现有表头，名称和年度不写入。";return;}
            if(!view.HeadersEnabled.Checked)return;
            if(companyHeader==null||yearSource==null||yearTarget==null)throw new InvalidOperationException("启用表头后，请完成名称目标、年度来源和年度目标。");
            if(companyHeader.Sheet!=templateSheet||yearTarget.Sheet!=templateSheet)throw new InvalidOperationException("表头目标必须与已编制区域同表。");
            int width=areas.Max(a=>a.Column+a.Columns)-areas.Min(a=>a.Column);
            if((yearSource.Rows!=1&&yearSource.Columns!=1)||Math.Max(yearSource.Rows,yearSource.Columns)!=width)throw new InvalidOperationException("年度来源有 "+Math.Max(yearSource.Rows,yearSource.Columns)+" 格，已编制区域跨度为 "+width+" 列，须相等且为单行或单列。");
            dynamic years=Sheet(yearSource.Sheet,bag),sheet=Sheet(templateSheet,bag),leadTable=Sheet(leadSheet,bag);var values=new List<string>();var formats=new List<string>();
            for(int c=0;c<width;c++)using(var one=new ComBag())
            {
                dynamic cell=one.Own(years.Range[CellAddress.Name(yearSource.Row+(yearSource.Rows>1?c:0),yearSource.Column+(yearSource.Rows==1?c:0))]);
                object merged=cell.MergeCells;if(merged==null||merged==DBNull.Value||Bool(merged))throw new InvalidOperationException("年度来源不能含合并格。");
                string text=Convert.ToString(cell.Text);if(string.IsNullOrWhiteSpace(text)||text.Trim('#').Length==0||SummaryAreaEngine.IsError((object)cell.Value2))throw new InvalidOperationException("年度来源显示为空、错误或列宽不足。");values.Add(text);formats.Add(Convert.ToString(cell.NumberFormat));
            }
            view.Message("表头预检："+picked.Count+" 家，跨度 "+width+"；年度 "+yearSource.Sheet+"!"+SummaryAreaEngine.Address(yearSource)+"："+string.Join("、",values.ToArray()));
            var leadRows=SummaryAreaEngine.LeadRows(LeadSelections());
            for(int i=0;i<picked.Count;i++)using(var one=new ComBag())
            {
                long start=(long)companyHeader.Column+(long)i*width,yearStart=(long)yearTarget.Column+(long)i*width;
                if(start+(view.RepeatNames.Checked?width:1)-1>16384||yearStart+width-1>16384)throw new InvalidOperationException("表头映射超出列边界。");
                var binding=leadRows.Single(x=>x.Row==picked[i].Row);dynamic source=one.Own(leadTable.Range[CellAddress.Name(binding.Row,binding.CodeColumn)]);
                string nameText=view.FormulaHeaders.Checked?Convert.ToString(source.Text):picked[i].Code;
                if(string.IsNullOrWhiteSpace(nameText)||nameText.Trim('#').Length==0)throw new InvalidOperationException("公司编码简称显示为空或列宽不足。");
                for(int c=0;c<(view.RepeatNames.Checked?width:1);c++)
                {
                    dynamic cell=one.Own(sheet.Range[CellAddress.Name(companyHeader.Row,(int)start+c)]);var part=new SummaryAreaPlan{Sheet=templateSheet,Row=companyHeader.Row,Column=(int)start+c,Rows=1,Columns=1};
                    if(Bool((object)cell.MergeCells)){dynamic merge=one.Own(cell.MergeArea);part=Coordinates(merge,one);part.Sheet=templateSheet;}
                    int group=paste.Column+i*width;if(part.Column<group||(long)part.Column+part.Columns>group+width)throw new InvalidOperationException("公司名称合并区域跨组："+SummaryAreaEngine.Address(part));
                    if(headerTargets.Any(x=>x.Row==part.Row&&x.Column==part.Column))continue;
                    if(!view.FormulaHeaders.Checked){dynamic range=one.Own(sheet.Range[SummaryAreaEngine.Address(part)]);RejectArraySpill(range,"公司名称目标",one);}
                    AddHeader(part,nameText);
                    if(view.FormulaHeaders.Checked){headerExpressions[HeaderKey(part)]=SummaryHeaderFormula.Cell(leadSheet,binding.Row,binding.CodeColumn);headerFormats[HeaderKey(part)]=Convert.ToString(source.NumberFormat);}
                }
                for(int c=0;c<width;c++)
                {
                    var part=new SummaryAreaPlan{Sheet=templateSheet,Row=yearTarget.Row,Column=(int)yearStart+c,Rows=1,Columns=1};dynamic cell=one.Own(sheet.Range[SummaryAreaEngine.Address(part)]);
                    if(Bool((object)cell.MergeCells))throw new InvalidOperationException("年度目标不能合并。");if(!view.FormulaHeaders.Checked)RejectSpecial(cell,"年度目标",one);
                    AddHeader(part,values[c]);if(view.FormulaHeaders.Checked){headerFormats[HeaderKey(part)]=formats[c];headerExpressions[HeaderKey(part)]=SummaryHeaderFormula.Cell(yearSource.Sheet,yearSource.Row+(yearSource.Rows>1?c:0),yearSource.Column+(yearSource.Rows==1?c:0));}
                }
                if(view.FormulaHeaders.Checked)yearHeaderArrays.Add(new SummaryHeaderArray{Area=new SummaryAreaPlan{Sheet=templateSheet,Row=yearTarget.Row,Column=(int)yearStart,Rows=1,Columns=width},Formula=SummaryHeaderFormula.Years(yearSource)});
                view.Message((i+1)+". "+picked[i].Code+"，名称来源 "+leadSheet+"!"+CellAddress.Name(binding.Row,binding.CodeColumn)+"；名称目标 "+CellAddress.Name(companyHeader.Row,(int)start)+"；年度 "+CellAddress.Name(yearTarget.Row,(int)yearStart)+"，"+(view.FormulaHeaders.Checked?SummaryHeaderFormula.Years(yearSource):"显示文本"));
            }
            if(view.FormulaHeaders.Checked)PlanExistingHeaderArrays(sheet,leadTable,leadRows,values,formats,bag);
            foreach(var part in headerTargets)
            {
                dynamic first=bag.Own(sheet.Range[CellAddress.Name(part.Row,part.Column)]);view.Message("表头授权覆盖 "+SummaryAreaEngine.Address(part)+"；原内容="+Convert.ToString(first.Formula));
                if(Bool((object)first.HasFormula)&&!oldHeaderArrays.Any(a=>SummaryAreaEngine.Intersects(a.Area,part)))originalTargetFormulas.Add(new SummaryAreaCell{Row=part.Row,Column=part.Column,Text=Convert.ToString(first.Formula)});
            }
            view.HeaderPreview.Text=picked.Count+" 家，"+(view.RepeatNames.Checked?"每年度列重复名称":"每家公司名称一次")+"；"+(view.FormulaHeaders.Checked?"直接公式关联 Lead":"显示文本")+"；既有数组 "+oldHeaderArrays.Count+" 块。";
        }
        private void AddHeader(SummaryAreaPlan part,string value)
        {
            if(areas.Any(a=>SummaryAreaEngine.Intersects(a,part))||targets.Any(a=>SummaryAreaEngine.Intersects(a,part))||headerTargets.Any(a=>SummaryAreaEngine.Intersects(a,part))||(yearSource.Sheet==part.Sheet&&SummaryAreaEngine.Intersects(yearSource,part))||(leadSheet==part.Sheet&&LeadSelections().Any(a=>SummaryAreaEngine.Intersects(a,part))))throw new InvalidOperationException("表头与已编制区域、复制目标、其他表头或来源冲突："+SummaryAreaEngine.Address(part));
            headerTargets.Add(part);headerValues.Add(value);
        }
        private SummaryHeaderArray ExistingHeaderArray(dynamic cell,ComBag bag)
        {
            dynamic range=null;bool spill=false;
            try { object state=cell.HasSpill;if(FormulaFlag(state,"表头","HasSpill")==1){dynamic parent=bag.Own(cell.SpillParent);range=bag.Own(parent.SpillingToRange);spill=true;} }
            catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}
            catch(System.Runtime.InteropServices.COMException error){if(error.HResult!=unchecked((int)0x80020003)&&error.HResult!=unchecked((int)0x80020006))throw;}
            if(range==null&&Bool((object)cell.HasArray))range=bag.Own(cell.CurrentArray);
            if(range==null){RejectArraySpill(cell,"表头目标",bag);return null;}
            var area=Coordinates(range,bag);area.Sheet=templateSheet;if((long)area.Rows*area.Columns>SummaryAreaEngine.CellLimit)throw new InvalidOperationException("既有表头数组过大。");
            dynamic anchor=bag.Own(((dynamic)Sheet(templateSheet,bag)).Range[CellAddress.Name(area.Row,area.Column)]);
            return new SummaryHeaderArray{Area=area,Formula=Convert.ToString(anchor.Formula),Dynamic=spill};
        }
        private void PlanExistingHeaderArrays(dynamic sheet,dynamic leadTable,IList<SummaryLeadRow> leadRows,IList<string> years,IList<string> formats,ComBag bag)
        {
            Dictionary<string,List<SummaryLeadRow>> nameLookup=null;
            foreach(var part in headerTargets.ToList())using(var one=new ComBag())
            {
                dynamic cell=one.Own(sheet.Range[CellAddress.Name(part.Row,part.Column)]);var existing=ExistingHeaderArray(cell,one);if(existing==null)continue;
                if(!oldHeaderArrays.Any(a=>SummaryAreaEngine.Address(a.Area)==SummaryAreaEngine.Address(existing.Area)))oldHeaderArrays.Add(existing);
            }
            foreach(var existing in oldHeaderArrays)
            {
                var area=existing.Area;view.Message("既有表头数组将整体重写并纳入备份："+SummaryAreaEngine.Address(area)+"；原公式="+existing.Formula);
                if(!view.ReplaceHeaderArrays.Checked)throw new InvalidOperationException("目标属于既有数组 "+SummaryAreaEngine.Address(area)+"；请核对日志并勾选允许整体重写既有表头数组，再预检。");
                if(area.Rows!=1||(area.Row!=companyHeader.Row&&area.Row!=yearTarget.Row))throw new InvalidOperationException("只允许重写名称行或年度行上的单行表头数组。");
                if(areas.Any(a=>SummaryAreaEngine.Intersects(a,area))||targets.Any(a=>SummaryAreaEngine.Intersects(a,area))||(yearSource.Sheet==templateSheet&&SummaryAreaEngine.Intersects(yearSource,area))||(leadSheet==templateSheet&&LeadSelections().Any(a=>SummaryAreaEngine.Intersects(a,area))))throw new InvalidOperationException("完整旧数组涉及已编制区域、数据目标或来源，不能扩大覆盖："+SummaryAreaEngine.Address(area));
                for(int col=area.Column;col<area.Column+area.Columns;col++)
                {
                    var part=new SummaryAreaPlan{Sheet=templateSheet,Row=area.Row,Column=col,Rows=1,Columns=1};if(headerTargets.Any(p=>SummaryAreaEngine.Intersects(p,part)))continue;
                    dynamic cell=bag.Own(sheet.Range[CellAddress.Name(area.Row,col)]);string text=Convert.ToString(cell.Text);if(SummaryAreaEngine.IsError((object)cell.Value2))throw new InvalidOperationException("旧数组目标外位置 "+CellAddress.Name(area.Row,col)+" 已有错误，无法安全恢复其 Lead 映射；请先修正或移除该旧数组。");
                    string expression=null,format=Convert.ToString(cell.NumberFormat);
                    if(!string.IsNullOrEmpty(text))
                    {
                        if(area.Row==companyHeader.Row)
                        {
                            if(nameLookup==null)
                            {
                                nameLookup=new Dictionary<string,List<SummaryLeadRow>>(StringComparer.Ordinal);
                                foreach(var row in leadRows)using(var one=new ComBag()){dynamic source=one.Own(leadTable.Range[CellAddress.Name(row.Row,row.CodeColumn)]);string display=Convert.ToString(source.Text);if(!nameLookup.ContainsKey(display))nameLookup[display]=new List<SummaryLeadRow>();nameLookup[display].Add(row);}
                            }
                            List<SummaryLeadRow> matches;if(!nameLookup.TryGetValue(text,out matches)||matches.Count!=1)throw new InvalidOperationException("旧数组目标外名称 "+CellAddress.Name(area.Row,col)+" 无法唯一匹配 Lead 编码简称，请调整来源或先处理旧数组。");expression=SummaryHeaderFormula.Cell(leadSheet,matches[0].Row,matches[0].CodeColumn);
                        }
                        else {var positions=Enumerable.Range(0,years.Count).Where(i=>years[i]==text).ToList();if(positions.Count!=1)throw new InvalidOperationException("旧数组目标外年度无法唯一匹配来源："+CellAddress.Name(area.Row,col));int index=positions[0];expression=SummaryHeaderFormula.Cell(yearSource.Sheet,yearSource.Row+(yearSource.Rows>1?index:0),yearSource.Column+(yearSource.Rows==1?index:0));format=formats[index];}
                    }
                    AddHeader(part,text);if(expression!=null){headerExpressions[HeaderKey(part)]=expression;headerFormats[HeaderKey(part)]=format;}view.Message("扩展授权 "+CellAddress.Name(part.Row,part.Column)+"：保留原显示并重新关联 Lead；"+(expression??"原为空"));
                }
            }
        }
        private void ClearHeaderArray(dynamic sheet,SummaryHeaderArray array,ComBag bag)
        {
            dynamic range=bag.Own(sheet.Range[SummaryAreaEngine.Address(array.Area)]),anchor=bag.Own(sheet.Range[CellAddress.Name(array.Area.Row,array.Area.Column)]);
            if(array.Dynamic)anchor.ClearContents();else range.ClearContents();
        }
        private void WriteHeaders(dynamic sheet,ComBag bag)
        {
            headersTouched=headerTargets.Count>0;
            foreach(var array in oldHeaderArrays)ClearHeaderArray(sheet,array,bag);
            for(int i=0;i<headerTargets.Count;i++)
            {
                var part=headerTargets[i];dynamic cell=bag.Own(sheet.Range[CellAddress.Name(part.Row,part.Column)]);string key=HeaderKey(part);
                if(view.FormulaHeaders.Checked&&headerExpressions.ContainsKey(key))
                {
                    cell.NumberFormat=headerFormats[key];if(!yearHeaderArrays.Any(a=>SummaryAreaEngine.Intersects(a.Area,part)))cell.Formula=headerExpressions[key];
                }
                else {object format=cell.NumberFormat;try{cell.NumberFormat="@";cell.Value2=headerValues[i];}finally{cell.NumberFormat=format;}}
            }
            foreach(var array in yearHeaderArrays){dynamic range=bag.Own(sheet.Range[SummaryAreaEngine.Address(array.Area)]);range.FormulaArray=array.Formula;range.Calculate();}
            for(int i=0;i<headerTargets.Count;i++)
            {
                var part=headerTargets[i];dynamic cell=bag.Own(sheet.Range[CellAddress.Name(part.Row,part.Column)]);string key=HeaderKey(part);bool formula=view.FormulaHeaders.Checked&&headerExpressions.ContainsKey(key);if(formula)cell.Calculate();
                string actual=formula?Convert.ToString(cell.Text):Convert.ToString(cell.Value2);
                if(Bool((object)cell.HasFormula)!=formula||actual!=headerValues[i])throw new InvalidOperationException("表头公式/显示值读回失败："+SummaryAreaEngine.Address(part)+"；期望="+headerValues[i]+"；实际="+actual);
            }
            if(headerTargets.Count>0)view.Message("表头写入及读回成功 "+headerTargets.Count+" 项；直接关联 Lead，不使用抵消表计数。未保存。");
        }
        private void VerifyHeaderArrayBounds(dynamic sheet,SummaryHeaderArray expected,ComBag bag)
        {
            dynamic anchor=bag.Own(sheet.Range[CellAddress.Name(expected.Area.Row,expected.Area.Column)]),range=null;
            if(expected.Dynamic)
            {
                if(FormulaFlag((object)anchor.HasSpill,"原表头数组","HasSpill")!=1)throw new InvalidOperationException("原动态数组未恢复为完整溢出范围；请保留备份并检查原来源参数。");
                range=bag.Own(anchor.SpillingToRange);
            }
            else
            {
                if(!Bool((object)anchor.HasArray))throw new InvalidOperationException("原表头数组未完整恢复。");range=bag.Own(anchor.CurrentArray);
            }
            var actual=Coordinates(range,bag);if(actual.Row!=expected.Area.Row||actual.Column!=expected.Area.Column||actual.Rows!=expected.Area.Rows||actual.Columns!=expected.Area.Columns)throw new InvalidOperationException("原表头数组范围发生变化："+SummaryAreaEngine.Address(expected.Area)+"；请先恢复原来源参数，备份保留。");
        }
        private void RestoreHeaders(dynamic sheet,dynamic original,ComBag bag)
        {
            if(!headersTouched)return;
            // Recalculate only the owned backup before restoring a dynamic formula:
            // changed source counts must not expand the restored spill outside consent.
            foreach(var array in oldHeaderArrays){dynamic source=bag.Own(original.Range[SummaryAreaEngine.Address(array.Area)]);source.Calculate();VerifyHeaderArrayBounds(original,array,bag);}
            foreach(var array in yearHeaderArrays){dynamic range=bag.Own(sheet.Range[SummaryAreaEngine.Address(array.Area)]);range.ClearContents();}
            foreach(var array in oldHeaderArrays)ClearHeaderArray(sheet,array,bag);
            foreach(var part in headerTargets)
            {
                if(oldHeaderArrays.Any(a=>SummaryAreaEngine.Intersects(a.Area,part)))continue;
                string address=SummaryAreaEngine.Address(part);dynamic source=bag.Own(original.Range[address]),destination=bag.Own(sheet.Range[address]);source.Copy(destination);
                foreach(var prior in originalTargetFormulas.Where(p=>p.Row>=part.Row&&p.Row<part.Row+part.Rows&&p.Column>=part.Column&&p.Column<part.Column+part.Columns)){dynamic cell=bag.Own(sheet.Range[CellAddress.Name(prior.Row,prior.Column)]);if(Convert.ToString(cell.Formula)!=prior.Text)cell.Formula=prior.Text;if(Convert.ToString(cell.Formula)!=prior.Text)throw new InvalidOperationException("表头原公式恢复失败。");}
            }
            foreach(var array in oldHeaderArrays)
            {
                string address=SummaryAreaEngine.Address(array.Area);dynamic source=bag.Own(original.Range[address]),destination=bag.Own(sheet.Range[address]);source.Copy(destination);dynamic anchor=bag.Own(sheet.Range[CellAddress.Name(array.Area.Row,array.Area.Column)]);
                if(Convert.ToString(anchor.Formula)!=array.Formula)throw new InvalidOperationException("旧表头数组原公式恢复核验失败："+address+"；备份保留。");VerifyHeaderArrayBounds(sheet,array,bag);
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
