using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // =====================================================================
    // 合并抵消（往来抵消 / 收入成本抵消）—— 功能区接线与业务编排
    //
    // 结构同 Connect.MergeHeaders.cs：窗体只做参数收集与结果呈现，一切 COM
    // 都在 Ribbon 宿主线程执行（窗体事件与回调同线程，不开后台 STA）。
    // 非模态显示取宿主主窗口作 owner——TrackerForm 锁死事故后全仓库统一
    // 写法，禁止 TopMost 无主窗。
    //
    // 两个功能区按钮共用一个窗体槽位：
    //   · 同 kind 已开 → 激活；
    //   · 异 kind 且旧窗空闲（不忙、无未确认批次）→ 换窗（旧窗资源先释放）；
    //   · 旧窗忙或本批未确认 → 提示先处理，不换窗。
    //
    // 写入目标 = 来源工作簿内的指定工作表或新建「抵消分录」表；写入前备份、
    // 快照，「取消操作」还原（新建表整表删除 / 已存在表快照回写）。
    // =====================================================================

    public partial class Connect
    {
        private static EliminationForm _eliminationForm;
        private static EliminationGateway _eliminationGateway;
        private static EliminationTargetHandle _eliminationTarget;
        private static MergeSnapshot _eliminationSnapshot;
        // 往来抵消会同时产出「差异表」与「结果表」两张表，各自维护目标句柄与快照
        private static EliminationTargetHandle _eliminationDiffTarget;
        private static MergeSnapshot _eliminationDiffSnapshot;
        // 预检产出的往来抵消结果（供预览网格与执行写入共用）
        private static EliminationDiffTable _eliminationDiff;
        private static string _eliminationBackupPath;
        // 上次建列映射网格的来源键（路径|工作表|表头行）；相同则保留用户改过的角色。
        private static string _eliminationHeadersKey = "";

        public IPictureDisp GetImage_menuElimination(object control)
        { return GetButtonImageCore("menuElimination"); }

        public IPictureDisp GetImage_btnEliminationReceivable(object control)
        { return GetButtonImageCore("btnEliminationReceivable"); }

        public IPictureDisp GetImage_btnEliminationRevenueCost(object control)
        { return GetButtonImageCore("btnEliminationRevenueCost"); }

        public void OnEliminationReceivable(object control)
        { OpenElimination(EliminationKind.Receivable); }

        public void OnEliminationRevenueCost(object control)
        { OpenElimination(EliminationKind.RevenueCost); }

        /// <summary>懒清理列映射临时记忆：清除已关闭工作簿的记忆条目（打开弹窗时执行）。</summary>
        private static void CleanupStaleMemory()
        {
            try
            {
                if (_app == null) return;
                dynamic app = _app;
                HashSet<string> open = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                dynamic books = app.Workbooks;
                int count = Convert.ToInt32(books.Count);
                for (int i = 1; i <= count; i++)
                {
                    string path = Convert.ToString(books.Item(i).FullName);
                    if (!string.IsNullOrEmpty(path)) open.Add(path);
                }
                EliminationForm.EliminationMapStore.CleanupExcept(open);
            }
            catch { }
        }

        /// <summary>
        /// 打开/切换抵消窗体。两个按钮共用一个窗体槽位（同 kind 激活、异 kind
        /// 空闲换窗、忙/未确认则提示）。
        /// </summary>
        private static void OpenElimination(EliminationKind kind)
        {
            if (_busy) return;
            EliminationForm current = _eliminationForm;
            if (current != null && !current.IsDisposed)
            {
                if (current.Kind == kind) { try { current.Activate(); } catch { } return; }
                if (current.IsBusy || current.CanUndo)
                {
                    UiDialog.Show("「" + current.Text + "」正在处理，或本批写入尚未「确认操作 / 取消操作」。\r\n请先处理完再打开另一种抵消。",
                        "合并抵消", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                // 旧窗空闲：先释放它持有的工作簿引用与写入态，再摘槽位关窗。
                // FormClosed 里的令牌判断此时不命中（槽位已空），只会 Dispose 旧窗，
                // 不会碰随后新建窗体的网关。
                ReleaseEliminationState();
                _eliminationForm = null;
                try { current.Close(); } catch { }
            }
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。");
                _eliminationHeadersKey = "";
                CleanupStaleMemory();
                EliminationForm form = new EliminationForm(kind);
                ApplyActiveSource(form);
                form.PreviewRequested += delegate { RunEliminationPreview(form); };
                form.ExecuteRequested += delegate { RunEliminationExecute(form); };
                form.UndoRequested += delegate { RunEliminationUndo(form); };
                form.ConfirmRequested += delegate { RunEliminationConfirm(form); };
                form.SheetsRequested += delegate(string path) { RunLoadEliminationSheets(form, path); };
                form.FormClosed += delegate
                {
                    bool mine = ReferenceEquals(_eliminationForm, form);
                    if (mine)
                    {
                        _eliminationForm = null;
                        ReleaseEliminationState();
                    }
                    try { form.PersistMapping(); } catch { }
                    form.Dispose();
                };
                _eliminationForm = form;
                // 与链接助手/表头匹配合并等非模态弹窗完全一致：取宿主主窗口作 owner 显示。
                IntPtr owner = GetHostWindowHandle();
                Log("合并抵消：打开前 enabled=" + HostDialogService.HostEnabled()
                    + " / 前台=" + NativeOle.GetForegroundWindow() + " / owner=" + owner);
                if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
                Log("合并抵消：Show 后 enabled=" + HostDialogService.HostEnabled()
                    + " / 前台=" + NativeOle.GetForegroundWindow()
                    + " / 弹窗=" + (form.IsHandleCreated ? form.Handle : IntPtr.Zero));
            }
            catch (Exception ex)
            {
                Log("合并抵消：" + ex);
                UiDialog.Show(ex.Message, "合并抵消", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>卸载/关文档时关掉还开着的抵消窗，避免挂在已死的宿主上。</summary>
        internal static void CloseEliminationForm()
        {
            EliminationForm form = _eliminationForm;
            _eliminationForm = null;
            if (form != null)
            {
                try
                {
                    if (!form.IsDisposed && form.IsHandleCreated)
                        form.BeginInvoke((MethodInvoker)delegate { try { form.Close(); } catch { } });
                }
                catch { try { form.Close(); } catch { } }
            }
            // 宿主即将卸载：BeginInvoke 未必还会执行，同步释放 COM 引用（幂等）。
            ReleaseEliminationState();
        }

        /// <summary>释放抵消功能持有的网关与写入态（关窗/换窗/卸载共用；幂等）。</summary>
        private static void ReleaseEliminationState()
        {
            EliminationGateway gateway = _eliminationGateway;
            _eliminationGateway = null;
            _eliminationTarget = null;
            _eliminationSnapshot = null;
            _eliminationDiffTarget = null;
            _eliminationDiffSnapshot = null;
            _eliminationDiff = null;
            _eliminationBackupPath = null;
            _eliminationHeadersKey = "";
            if (gateway != null)
            {
                try { gateway.Dispose(); }
                catch (Exception ex) { Log("合并抵消：释放工作簿引用失败 " + ex.Message); }
            }
        }

        private static EliminationGateway Gateway()
        {
            if (_eliminationGateway == null) _eliminationGateway = new EliminationGateway(_app);
            return _eliminationGateway;
        }

        // ---------------------------- 来源 ----------------------------

        /// <summary>
        /// 把当前活动工作簿与活动工作表填进来源区。COM 只能在 Ribbon 宿主线程做，
        /// 因此放在这里而不是弹窗构造函数里。不满足条件时留空并说明理由。
        /// </summary>
        private static void ApplyActiveSource(EliminationForm form)
        {
            try
            {
                using (ComBag bag = new ComBag())
                {
                    object book = null;
                    try { book = bag.Own(((dynamic)_app).ActiveWorkbook); } catch { }
                    if (book == null) return;
                    string path = Convert.ToString(((dynamic)book).FullName);
                    string sheet = "";
                    try
                    {
                        object active = bag.Own(((dynamic)_app).ActiveSheet);
                        sheet = Convert.ToString(((dynamic)active).Name);
                    }
                    catch { }
                    if (string.IsNullOrEmpty(path))
                    {
                        form.AppendLog("当前工作簿尚未保存到磁盘，请保存后重开本窗口，或点「选择来源…」指定一个 .xlsx 工作簿。");
                        return;
                    }
                    if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        form.AppendLog("当前工作簿不是 .xlsx（合并抵消只处理 .xlsx）：" + path + "。请点「选择来源…」另选。");
                        return;
                    }
                    form.SetSource(path, sheet);
                }
            }
            catch (Exception ex) { Log("合并抵消读取活动工作簿：" + ex); }
        }

        /// <summary>只为填下拉而读一次工作表名：已打开则复用，否则只读打开。</summary>
        private static void RunLoadEliminationSheets(EliminationForm form, string path)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。");
                if (string.IsNullOrEmpty(path)) throw new InvalidOperationException("路径为空。");
                if (!File.Exists(path)) throw new InvalidOperationException("文件不存在：" + path);
                List<string> names = Gateway().SheetNames(path);
                form.SetSheets(names, null);
            }
            catch (Exception ex)
            {
                Log("合并抵消读取工作表名：" + ex);
                form.SetSheets(null, ex.Message);
            }
            finally { _busy = false; }
        }

        // ---------------------------- 预检 ----------------------------

        private static void RunEliminationPreview(EliminationForm form)
        {
            if (_busy) return;
            _busy = true;
            form.SetBusy(true);
            try
            {
                EliminationRequest request = form.BuildRequest();
                string invalid = ValidateEliminationRequest(request);
                if (invalid != null) { form.SetLog(invalid); return; }

                EliminationGateway gateway = Gateway();

                form.SetStatus("正在读取来源表头…");
                MergeHeaderSet headers = gateway.ReadHeaders(request.Path, request.SheetName, request.HeaderRow);
                string headerError = headers.Validate("来源工作簿");
                if (headerError != null) { form.SetHeaders(null, 0, headerError); return; }
                int dataRows = headers.LastUsedRow - request.HeaderRow;
                if (dataRows < 0) dataRows = 0;

                // 同一来源重复预检时保留用户在列映射网格里改过的角色；
                // 换了来源/工作表/表头行/业务种类（或网格被清空）才按表头重建。
                // 键里必须含 kind：两种业务的角色集合不同，共用缓存会串。
                string key = request.Kind + "|" + request.Path + "|" + request.SheetName + "|"
                    + request.HeaderRow.ToString(CultureInfo.InvariantCulture);
                if (form.MapRowCount == 0 || key != _eliminationHeadersKey)
                {
                    form.SetHeaders(BuildColumns(headers, request.Kind), dataRows, null);
                    _eliminationHeadersKey = key;
                    request = form.BuildRequest();
                    if (request.Error != null) { form.SetLog(request.Error); return; }
                }

                string mappingError = EliminationEngine.ValidateMapping(request.Kind, request.Columns);
                if (mappingError != null)
                {
                    form.SetLog(mappingError + "\r\n请在列映射网格中修改角色后重新点「读取表头并预检」；"
                        + "映射修改完成前「执行写入」不可用。");
                    return;
                }

                if (request.Kind == EliminationKind.RevenueCost)
                {
                    // 收入成本抵消走矩阵式底稿
                    string matrixError = EliminationEngine.ValidateMatrixMapping(request.Kind, request.Columns);
                    if (matrixError != null)
                    {
                        form.SetLog(matrixError + "\r\n请在列映射网格中修改角色后重新点「读取表头并预检」；"
                            + "映射修改完成前「执行写入」不可用。");
                        return;
                    }

                    form.SetStatus("正在读取源数据…");
                    List<EliminationRow> matrixRows = gateway.ReadRows(
                        request.Path, request.SheetName, request.HeaderRow, request.Columns,
                        headers.LastUsedRow, false);
                    MatrixOptions options = BuildMatrixOptions(request, headers);
                    EliminationMatrix matrix = EliminationMatrixBuilder.Build(matrixRows, request.Columns, options);
                    form.SetPreviewMatrix(matrix);
                    if (matrix.Errors.Count == 0 && matrix.Rows.Count > 0)
                        EliminationForm.EliminationMapStore.Save(
                            EliminationForm.EliminationMapStore.Key(request.Kind, request.Path), request.Columns);
                    Log("合并抵消预检：收入成本抵消（矩阵底稿） 来源 " + request.Path + " [" + request.SheetName
                        + "] 交易对 " + matrix.GroupCount + " 组 / 年度 " + matrix.Years.Count + " 个");
                    return;
                }

                // 往来抵消（Receivable）：走「差异表 + 结果表」两张表的矩阵式流程
                form.SetStatus("正在读取源数据…");
                List<EliminationRow> rows = gateway.ReadRows(
                    request.Path, request.SheetName, request.HeaderRow, request.Columns, headers.LastUsedRow, false);
                MatrixOptions wlOptions = BuildMatrixOptions(request, headers);
                EliminationDiffTable diff = null;
                EliminationMatrix wlResult = null;
                EliminationDiffBuilder.Build(rows, request.Columns, wlOptions, out diff, out wlResult);
                _eliminationDiff = diff;
                // 预览网格展示结果表（矩阵式，与收入成本一致的观感）；差异表统计写进日志
                form.SetPreviewMatrix(wlResult);
                if (diff != null)
                    form.AppendLog("往来抵消差异表：" + diff.Rows.Count + " 行 = 主体对 " + diff.GroupCount
                        + " 组 × 年度 " + diff.Years.Count + " 个 × 科目对 " + diff.RowsPerYear + "；"
                        + "条件列筛掉 " + diff.FilteredRows + " 行、跳过 " + diff.SkippedRows + " 行。");
                if (diff != null && diff.Errors.Count == 0 && diff.Rows.Count > 0)
                    EliminationForm.EliminationMapStore.Save(
                        EliminationForm.EliminationMapStore.Key(request.Kind, request.Path), request.Columns);
                Log("合并抵消预检：往来抵消 来源 " + request.Path + " [" + request.SheetName + "] 主体对 "
                    + (diff == null ? 0 : diff.GroupCount) + " 组 / 差异行 "
                    + (diff == null ? 0 : diff.Rows.Count) + " 行 / 结果行 "
                    + (wlResult == null ? 0 : wlResult.Rows.Count) + " 行");
            }
            catch (Exception ex)
            {
                Log("合并抵消预检：" + ex);
                form.SetLog("预检失败：" + ex.Message);
            }
            finally
            {
                form.SetBusy(false);
                _busy = false;
            }
        }

        /// <summary>表头快照 → 列映射（角色按表头文本猜，用户可在网格覆盖）。</summary>
        private static List<EliminationColumn> BuildColumns(MergeHeaderSet headers, EliminationKind kind)
        {
            List<EliminationColumn> columns = new List<EliminationColumn>();
            foreach (MergeColumn source in headers.Columns)
            {
                columns.Add(new EliminationColumn
                {
                    Number = source.Number,
                    Letter = source.Letter,
                    Header = source.Text,
                    Role = EliminationEngine.GuessRole(source.Text, kind)
                });
            }
            return columns;
        }

        private static bool HasRole(List<EliminationColumn> columns, EliminationRole role)
        {
            foreach (EliminationColumn column in columns)
                if (column.Role == role) return true;
            return false;
        }

        /// <summary>界面层之外的请求校验（路径/工作表/表头行）；null 表示通过。</summary>
        private static string ValidateEliminationRequest(EliminationRequest request)
        {
            if (request.Error != null) return request.Error;
            if (request.Path.Length == 0) return "请先选择来源工作簿。";
            if (request.SheetName.Length == 0) return "请选择来源工作表。";
            if (!File.Exists(request.Path)) return "来源文件不存在：" + request.Path;
            if (!string.Equals(Path.GetExtension(request.Path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                return "来源工作簿需为 .xlsx：" + request.Path;
            return null;
        }

        // ---------------------------- 执行 ----------------------------

        private static void RunEliminationExecute(EliminationForm form)
        {
            if (_busy) return;
            _busy = true;
            form.SetBusy(true);
            try
            {
                EliminationRequest request = form.BuildRequest();
                string invalid = ValidateEliminationRequest(request);
                if (invalid != null) { form.SetLog(invalid); return; }

                int startRow, startColumn;
                if (!EliminationGateway.TryParseCell(request.StartCell, out startRow, out startColumn))
                {
                    form.SetLog("起始单元格无法解析：" + request.StartCell + "（示例：A1、B3）。");
                    return;
                }

                EliminationGateway gateway = Gateway();

                // 执行时以当前界面状态重算，不依赖上次预检缓存（与表头匹配合并一致）。
                form.SetStatus("正在计算分录…");
                MergeHeaderSet headers = gateway.ReadHeaders(request.Path, request.SheetName, request.HeaderRow);
                string headerError = headers.Validate("来源工作簿");
                if (headerError != null) throw new InvalidOperationException(headerError);
                string mappingError = EliminationEngine.ValidateMapping(request.Kind, request.Columns);
                if (mappingError != null) { form.SetLog(mappingError); return; }

                // 两种业务的写入计划都走矩阵式（BuildMatrixPlan），样式链也统一走矩阵样式
                // （MergeMatrixHeader/StyleMatrixHeader/FinishMatrixOutput）。
                // 旧代码此处按 kind 分流，往来抵消走了「分录式」else 样式（StyleEntryHeader），
                // 导致结果表无合并/无填充/无边框/无千分位（实机验收发现）。
                bool receivable = request.Kind == EliminationKind.Receivable;
                EliminationMatrix built = null;
                EliminationResult result = null;
                MergeWritePlan plan;
                // 往来抵消需要「先建差异表取得真实表名 → 再重算」（结果表金额是跨表引用）
                MatrixOptions wlOptions = null;
                List<EliminationRow> wlRows = null;

                if (!receivable)
                {
                    string matrixError2 = EliminationEngine.ValidateMatrixMapping(request.Kind, request.Columns);
                    if (matrixError2 != null) { form.SetLog(matrixError2); return; }

                    // 起始位置必须在 Build 前定好：公式锚点据此生成，之后不能再改。
                    MatrixOptions options = BuildMatrixOptions(request, headers);
                    options.StartRow = startRow;
                    options.StartColumn = startColumn;
                    List<EliminationRow> matrixRows = gateway.ReadRows(
                        request.Path, request.SheetName, request.HeaderRow, request.Columns,
                        headers.LastUsedRow, false);
                    built = EliminationMatrixBuilder.Build(matrixRows, request.Columns, options);
                    if (built.Rows.Count == 0) { form.SetLog("没有可写入的内容。"); return; }
                    if (built.Errors.Count > 0)
                    {
                        form.SetLog(string.Join("\r\n", built.Errors.ToArray()));
                        return;
                    }
                    plan = EliminationGateway.BuildMatrixPlan(built, options);
                }
                else
                {
                    // 往来抵消：同时产出「差异表」与「结果表」两张表（矩阵式）。
                    // 确认前用默认差异表名粗算一次供确认文案；确认后建差异表拿到真实表名再重算
                    // （结果表金额 = 跨表引用差异表的抵消金额格，表名重名自动编号时必须用实际名）。
                    wlOptions = BuildMatrixOptions(request, headers);
                    wlOptions.StartRow = startRow;
                    wlOptions.StartColumn = startColumn;
                    wlRows = gateway.ReadRows(
                        request.Path, request.SheetName, request.HeaderRow, request.Columns,
                        headers.LastUsedRow, false);
                    EliminationDiffTable diff = null;
                    EliminationDiffBuilder.Build(wlRows, request.Columns, wlOptions, out diff, out built);
                    _eliminationDiff = diff;
                    if (diff == null || diff.Rows.Count == 0) { form.SetLog("没有可写入的内容。"); return; }
                    if (diff.Errors.Count > 0)
                    {
                        form.SetLog(string.Join("\r\n", diff.Errors.ToArray()));
                        return;
                    }
                    plan = EliminationGateway.BuildMatrixPlan(built, wlOptions);
                }

                string guard = MergePlanner.Guard(plan);
                if (guard != null) { form.SetLog(guard); return; }

                string defaultName = EliminationGateway.MatrixSheetNameFor(request.Kind);
                string targetDesc = request.TargetSheetItem == EliminationGateway.NewSheetItem
                    ? "新建工作表（默认名「" + defaultName + "」，重名自动编号）"
                    : request.TargetSheetItem;
                bool sameSheet = string.Equals(request.TargetSheetItem, request.SheetName, StringComparison.Ordinal);

                if (UiDialog.Show(form, !receivable
                        ? ConfirmationText(request, built, plan, targetDesc, sameSheet)
                        : ConfirmationText(request, _eliminationDiff, built, plan, targetDesc, sameSheet),
                    "确认执行" + form.Text,
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2) != DialogResult.OK)
                {
                    form.AppendLog("已取消，未写入任何内容。");
                    return;
                }

                // 确认后才建目标：用户取消不留下空表。
                form.SetStatus("正在定位目标工作表…");
                // 往来抵消：先建「差异表」拿真实表名（重名自动编号），再据此重算两张表
                if (receivable)
                {
                    EliminationTargetHandle diffTarget = gateway.CreateTarget(
                        request.Path, request.DiffSheetItem, EliminationGateway.DiffDefaultSheetName);
                    _eliminationDiffTarget = diffTarget;
                    if (request.DiffSheetItem == EliminationGateway.NewSheetItem
                        && !string.Equals(diffTarget.SheetName, EliminationGateway.DiffDefaultSheetName, StringComparison.Ordinal))
                        form.AppendLog("差异表重名，已自动编号为「" + diffTarget.SheetName + "」。");
                    wlOptions.DiffSheetName = diffTarget.SheetName;
                    EliminationDiffTable diff2 = null;
                    EliminationMatrix rebuilt = null;
                    EliminationDiffBuilder.Build(wlRows, request.Columns, wlOptions, out diff2, out rebuilt);
                    _eliminationDiff = diff2;
                    built = rebuilt;
                    plan = EliminationGateway.BuildMatrixPlan(built, wlOptions);
                }
                EliminationTargetHandle target = gateway.CreateTarget(request.Path, request.TargetSheetItem, defaultName);
                _eliminationTarget = target;
                if (sameSheet && !target.IsNewSheet)
                    form.AppendLog("注意：目标工作表与来源相同，请确认写入区域未覆盖源数据。");

                form.SetStatus("正在保存并备份目标工作簿…");
                string saved = gateway.SaveTargetBookPath(target.Book);
                _eliminationBackupPath = gateway.CreateBackup(saved);

                // 快照先取，再清区域：这样「取消操作」能还原到写入前的原样。
                form.SetStatus("正在快照写入区域…");
                _eliminationSnapshot = gateway.Capture(target, plan);
                // 目标若之前写过矩阵，残留的合并单元格会让批量赋值失败，先拆掉。
                gateway.PrepareTargetRegion(target, plan);

                form.SetStatus("正在写入…");
                // 往来抵消：**先写差异表、后写结果表**——结果表的金额是跨表引用差异表，
                // 引用目标先有数据，写入核对也更稳（实机曾报「公式写入核对失败：G4」）。
                int diffCells = 0;
                if (receivable)
                {
                    MergeWritePlan diffPlan = EliminationGateway.BuildDiffPlan(_eliminationDiff, wlOptions);
                    string diffGuard = MergePlanner.Guard(diffPlan);
                    if (diffGuard != null) { form.SetLog(diffGuard); return; }
                    EliminationTargetHandle diffTarget = _eliminationDiffTarget;
                    _eliminationDiffSnapshot = gateway.Capture(diffTarget, diffPlan);
                    gateway.PrepareTargetRegion(diffTarget, diffPlan);
                    gateway.Write(diffTarget, diffPlan);
                    gateway.StyleDiffHeader(diffTarget, startRow, startColumn, _eliminationDiff.ColumnCount);
                    // 差异表：第 1 行表头、第 2 行起数据；金额列 G-J（索引 6-9）；序号 A/年度 B 居中
                    int diffDataTop = startRow + 1;
                    int diffDataRows = _eliminationDiff.Rows.Count > 0
                        ? (_eliminationDiff.Rows[_eliminationDiff.Rows.Count - 1].RowNumber - diffDataTop + 1) : 0;
                    int diffGap = wlOptions.GroupGap > 0 ? wlOptions.GroupGap : 0;
                    int diffGaps = diffGap > 0 && _eliminationDiff.GroupCount > 1
                        ? (_eliminationDiff.GroupCount - 1) * diffGap : 0;
                    gateway.FinishMatrixOutput(diffTarget, startRow, 1, diffDataTop, startColumn,
                        _eliminationDiff.ColumnCount, diffDataRows + diffGaps,
                        new int[] { 6, 7, 8, 9 }, new int[] { 0, 1 });
                    diffCells = diffPlan.TotalCells;
                    form.AppendLog("已写入 " + diffCells + " 格到「" + diffTarget.SheetName + "」（往来抵消差异）");
                }

                gateway.Write(target, plan);
                // 矩阵样式链两种业务恒走（旧代码按 kind 分流，往来走了分录式样式致结果表无格式）
                {
                    int firstHeaderRow = startRow + (request.TotalRow ? 1 : 0);
                    gateway.MergeMatrixHeader(target, built, startRow, startColumn, request.TotalRow);
                    gateway.StyleMatrixHeader(target, built, firstHeaderRow, startColumn, request.TotalRow);
                    // 接入 (2) 规范：金额列会计格式（千位分隔）、数据区对齐、列宽 AutoFit
                    int dataTop = firstHeaderRow + 2;
                    int dataRows = built.Rows.Count > 0
                        ? (built.Rows[built.Rows.Count - 1].RowNumber - dataTop + 1) : 0;
                    int[] moneyCols = EliminationGateway.MoneyColumnIndices(built);
                    int[] centerCols = new int[] { 3, 4 };   // 调整事项、会计科目
                    // 「格式参考」样张变体：宋体 11、slantDashDot 框线、SUBTOTAL 行也走会计格式
                    gateway.FinishMatrixOutput(target, firstHeaderRow, 2, dataTop, startColumn,
                        built.Columns.Count, dataRows, moneyCols, centerCols, true);
                }
                gateway.SaveTargetBook(target.Book);

                form.MarkWritten();
                form.AppendLog("已写入 " + plan.TotalCells + " 格到「" + target.SheetName + "」"
                    + (target.IsNewSheet ? "（本次新建）" : "") + "；备份 " + _eliminationBackupPath);
                form.SetStatus("写入完成；如需还原请点「取消操作」，确认无误请点「确认操作」。");
                Log("合并抵消写入：" + plan.TotalCells + " 格 → " + request.Path + " [" + target.SheetName + "]");
            }
            catch (Exception ex)
            {
                Log("合并抵消执行：" + ex);
                form.AppendLog("执行失败：" + ex.Message + "（目标工作表未改动或已按快照还原）");
                TryEliminationRollback(form);
                // 日志区不显眼；用模态框把真实失败原因直接端到用户面前，避免"以为执行了"。
                UiDialog.Show(form, "执行失败：" + ex.Message + "\r\n（目标工作表未改动或已按快照还原；详见日志。）",
                    form.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                form.SetBusy(false);
                _busy = false;
            }
        }

        /// <summary>界面选项 → 矩阵底稿的输出选项。</summary>
        private static MatrixOptions BuildMatrixOptions(EliminationRequest request, MergeHeaderSet headers)
        {
            MatrixOptions options = new MatrixOptions();
            options.SourceSheetName = request.SheetName;
            options.SourceBookName = "";
            options.ConditionValue = request.ConditionValue;
            options.HasConditionColumn = HasRole(request.Columns, EliminationRole.Condition);
            options.TotalRow = request.TotalRow;
            options.StaticValues = request.StaticValues;
            options.FullColumnRefs = request.FullColumnRefs;
            // 「引用整列」统一表达「自动扩展」：勾选 → SUMIFS 整列、现算 SUMPRODUCT 算到工作表末尾，
            // 源表新增行自动纳入；取消 → 两者都限定到当前数据区，快，但新增行后需重新执行。
            options.LimitComputedRange = !request.FullColumnRefs;
            options.GroupGap = request.GapRows > 0 ? request.GapRows : 0;
            options.SourceFirstRow = request.HeaderRow + 1;
            options.SourceLastRow = headers == null ? 0 : headers.LastUsedRow;
            options.StartRow = 1;
            options.StartColumn = 1;
            return options;
        }

        /// <summary>差异表与结果表共用一套输出选项（起始位置一致）。</summary>
        private static MatrixOptions DiffOptions(EliminationRequest request, MergeHeaderSet headers,
            int startRow, int startColumn)
        {
            MatrixOptions options = BuildMatrixOptions(request, headers);
            options.StartRow = startRow;
            options.StartColumn = startColumn;
            return options;
        }

        /// <summary>往来抵消的确认文案（差异表 + 结果表两张表，各自按下拉选择）。</summary>
        private static string ConfirmationText(EliminationRequest request, EliminationDiffTable diff,
            EliminationMatrix matrix, MergeWritePlan plan, string targetDesc, bool sameSheet)
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine("来源：" + request.Path + "  [" + request.SheetName + "]  表头行 "
                + request.HeaderRow.ToString(CultureInfo.InvariantCulture));
            string diffDesc = request.DiffSheetItem == EliminationGateway.NewSheetItem
                ? "新建工作表（默认名「" + EliminationGateway.DiffDefaultSheetName + "」，重名自动编号）"
                : request.DiffSheetItem;
            text.AppendLine("目标1（差异表）：[" + diffDesc + "]");
            text.AppendLine("目标2（结果表）：[" + targetDesc + "]  起始 " + request.StartCell);
            text.AppendLine();
            if (diff != null)
            {
                text.AppendLine("主体对 " + diff.GroupCount + " 组 × 年度 " + diff.Years.Count + " 个 × 科目对 "
                    + diff.RowsPerYear + " → 差异表 " + diff.Rows.Count + " 行（每主体对后空 "
                    + (request.GapRows > 0 ? request.GapRows : 0) + " 行）");
                text.AppendLine("年度：" + string.Join("、", diff.Years.ToArray()));
                if (diff.FilteredRows > 0) text.AppendLine("条件列筛掉 " + diff.FilteredRows + " 行。");
            }
            if (matrix != null)
                text.AppendLine("结果表 " + matrix.Rows.Count + " 行（仅输出抵消金额不为 0 的科目对）。");
            if (request.StaticValues)
                text.AppendLine("金额以静态数值写入（不写公式）。");
            else
                text.AppendLine("金额写为活公式（SUMPRODUCT 前缀匹配），源数据变动后结果自动重算。");
            if (sameSheet)
                text.AppendLine("⚠ 目标工作表与来源相同，请确认写入区域未覆盖源数据。");
            text.AppendLine();
            text.AppendLine("写入前会自动保存并备份目标工作簿（同目录「ExcelIterate合并备份」），写入后可用「取消操作」还原。是否继续？");
            return text.ToString();
        }

        /// <summary>矩阵式底稿的确认文案。</summary>
        private static string ConfirmationText(EliminationRequest request, EliminationMatrix matrix,
            MergeWritePlan plan, string targetDesc, bool sameSheet)
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine("来源：" + request.Path + "  [" + request.SheetName + "]  表头行 "
                + request.HeaderRow.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("目标：" + request.Path + "  [" + targetDesc + "]  起始 " + request.StartCell);
            text.AppendLine();
            text.AppendLine("交易对 " + matrix.GroupCount + " 组 × 年度 " + matrix.Years.Count + " 个 → "
                + matrix.Rows.Count + " 行、" + matrix.ColumnCount + " 列，共 " + plan.TotalCells + " 个单元格"
                + (matrix.Balanced ? "（借贷平衡）" : "（⚠ 借贷不平，请检查）"));
            text.AppendLine("年度：" + string.Join("、", matrix.Years.ToArray()));
            if (matrix.SkippedRows > 0)
                text.AppendLine("跳过 " + matrix.SkippedRows + " 行源数据（原因见预检报告）。");
            if (matrix.FilteredRows > 0)
                text.AppendLine("条件列筛掉 " + matrix.FilteredRows + " 行。");
            if (request.StaticValues)
                text.AppendLine("金额以静态数值写入（不写公式）。");
            else
                text.AppendLine("金额写为活公式（SUMIFS / SUMPRODUCT），源数据变动后结果自动重算。");
            if (sameSheet)
                text.AppendLine("⚠ 目标工作表与来源相同，请确认写入区域未覆盖源数据。");
            text.AppendLine();
            text.AppendLine("写入前会自动保存并备份目标工作簿（同目录「ExcelIterate合并备份」），写入后可用「取消操作」还原。是否继续？");
            return text.ToString();
        }

        private static string ConfirmationText(EliminationRequest request, EliminationResult result,
            MergeWritePlan plan, string targetDesc, bool sameSheet)
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine("来源：" + request.Path + "  [" + request.SheetName + "]  表头行 "
                + request.HeaderRow.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("目标：" + request.Path + "  [" + targetDesc + "]  起始 " + request.StartCell);
            text.AppendLine();
            text.AppendLine("分录 " + result.Entries.Count + " 条，共 " + plan.TotalCells + " 个单元格；总借 "
                + result.TotalDebit.ToString("0.00", CultureInfo.InvariantCulture)
                + "、总贷 " + result.TotalCredit.ToString("0.00", CultureInfo.InvariantCulture)
                + (result.Balanced ? "（借贷平衡）" : "（⚠ 借贷不平，请检查）"));
            if (result.SkippedRows > 0)
                text.AppendLine("跳过 " + result.SkippedRows + " 行源数据（原因见预检报告）。");
            if (sameSheet)
                text.AppendLine("⚠ 目标工作表与来源相同，请确认写入区域未覆盖源数据。");
            text.AppendLine();
            text.AppendLine("写入前会自动保存并备份目标工作簿（同目录「ExcelIterate合并备份」），写入后可用「取消操作」还原。是否继续？");
            return text.ToString();
        }

        private static void TryEliminationRollback(EliminationForm form)
        {
            EliminationTargetHandle target = _eliminationTarget;
            if (target == null || target.Sheet == null) return;
            // 没有快照说明还没动笔（Capture 之前失败），已存在的表无需还原。
            if (!target.IsNewSheet && _eliminationSnapshot == null) return;
            try
            {
                EliminationGateway gateway = Gateway();
                gateway.Restore(target, _eliminationSnapshot);
                // 往来抵消的差异表一并还原
                if (_eliminationDiffTarget != null && _eliminationDiffSnapshot != null)
                {
                    gateway.Restore(_eliminationDiffTarget, _eliminationDiffSnapshot);
                    form.AppendLog("已删除/还原「" + _eliminationDiffTarget.SheetName + "」（往来抵消差异）。");
                }
                try { gateway.SaveTargetBook(target.Book); } catch { }
                form.AppendLog(target.IsNewSheet
                    ? "已删除本次新建的工作表「" + target.SheetName + "」。"
                    : "已用写入前快照还原目标工作表。");
            }
            catch (Exception ex)
            {
                Log("合并抵消还原失败：" + ex);
                form.AppendLog("还原失败：" + ex.Message + "；请用备份文件手动恢复：" + _eliminationBackupPath);
            }
            finally
            {
                _eliminationTarget = null;
                _eliminationSnapshot = null;
                _eliminationDiffTarget = null;
                _eliminationDiffSnapshot = null;
            }
        }

        // ---------------------------- 取消 / 确认 ----------------------------

        private static void RunEliminationUndo(EliminationForm form)
        {
            if (_busy) return;
            _busy = true;
            form.SetBusy(true);
            try
            {
                EliminationTargetHandle target = _eliminationTarget;
                if (target == null || target.Sheet == null)
                {
                    form.AppendLog("没有可还原的批次。");
                    return;
                }
                if (UiDialog.Show(form, target.IsNewSheet
                    ? "将删除本次新建的工作表「" + target.SheetName + "」并自动保存。是否继续？"
                    : "将用写入前的快照还原目标工作表，并自动保存。是否继续？",
                    "取消操作", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2) != DialogResult.OK)
                {
                    form.AppendLog("已取消还原。");
                    return;
                }
                EliminationGateway gateway = Gateway();
                gateway.Restore(target, _eliminationSnapshot);
                gateway.SaveTargetBook(target.Book);
                form.MarkUndone();
                form.AppendLog(target.IsNewSheet
                    ? "已删除新建工作表「" + target.SheetName + "」。文件备份保留：" + _eliminationBackupPath
                    : "已还原到写入前状态。文件备份保留：" + _eliminationBackupPath);
                Log("合并抵消：已还原");
                _eliminationTarget = null;
                _eliminationSnapshot = null;
            }
            catch (Exception ex)
            {
                Log("合并抵消取消操作：" + ex);
                form.AppendLog("还原失败：" + ex.Message + "；请用备份文件手动恢复：" + _eliminationBackupPath);
            }
            finally
            {
                form.SetBusy(false);
                _busy = false;
            }
        }

        private static void RunEliminationConfirm(EliminationForm form)
        {
            _eliminationSnapshot = null;
            Log("合并抵消：已确认，本批结束。备份保留于 " + _eliminationBackupPath);
        }
    }
}
