using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace Excel_Iterate
{
    // =====================================================================
    // 跨工作簿表头匹配合并 —— 功能区接线与业务编排
    //
    // 所有 COM 调用都在 Ribbon 宿主线程（与 ColumnFilterForm 的 ApplyRequested
    // 同构）。回调必须写在 Connect*.cs 内，否则 configure_core.py 的回调
    // 存在性校验会失败。
    // =====================================================================

    public partial class Connect
    {
        private static MergeHeadersForm _mergeForm;
        private static MergeWorkbookGateway _mergeGateway;
        private static MergeWorkbookHandle _mergeTarget;
        private static MergeSnapshot _mergeSnapshot;
        private static string _mergeBackupPath;







        /// <summary>卸载/关文档时关掉还开着的合并窗，避免挂在已死的宿主上。</summary>
        internal static void CloseMergeForm()
        {
            MergeHeadersForm form = _mergeForm;
            _mergeForm = null;
            if (form == null) return;
            try
            {
                if (!form.IsDisposed && form.IsHandleCreated)
                    form.BeginInvoke((System.Windows.Forms.MethodInvoker)delegate { try { form.Close(); } catch { } });
            }
            catch { try { form.Close(); } catch { } }
        }

        /// <summary>
        /// 把当前活动工作簿与活动工作表填进目标区。COM 只能在 Ribbon 宿主线程做，
        /// 因此放在这里而不是弹窗构造函数里。不满足条件时留空并说明理由，避免出现
        /// “填了但预检必失败”的假可用状态。
        /// </summary>
        private static void ApplyActiveTarget(MergeHeadersForm form)
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
                        form.AppendLog("当前工作簿尚未保存到磁盘，请保存后重开本窗口，或点「选择目标…」指定一个 .xlsx 工作簿。");
                        return;
                    }
                    if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        form.AppendLog("当前工作簿不是 .xlsx（表头匹配只处理 .xlsx）：" + path + "。请点「选择目标…」另选。");
                        return;
                    }
                    form.SetTarget(path, sheet);
                    try { form.SetTargetSheets(MergeWorkbookGateway.SheetNames(book), null); }
                    catch (Exception ex) { Log("表头匹配合并读取活动工作表：" + ex); }
                }
            }
            catch (Exception ex) { Log("表头匹配合并读取活动工作簿：" + ex); }
        }

        /// <summary>
        /// 只为填下拉而读一次工作表名：已打开则复用，否则只读打开；读完全交给 Gateway
        /// 统一释放（这里不关，避免误关用户自己已打开的簿）。
        /// </summary>
        private static void RunLoadSheets(MergeHeadersForm form, string path, bool asSource)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。");
                if (string.IsNullOrEmpty(path)) throw new InvalidOperationException("路径为空。");
                if (!File.Exists(path)) throw new InvalidOperationException("文件不存在：" + path);
                if (_mergeGateway == null) _mergeGateway = new MergeWorkbookGateway(_app);
                MergeWorkbookHandle handle = _mergeGateway.Resolve(path, "", true);
                List<string> names = MergeWorkbookGateway.SheetNames(handle.Book);
                if (asSource) form.SetSourceSheets(path, names, null);
                else form.SetTargetSheets(names, null);
            }
            catch (Exception ex)
            {
                Log("表头匹配合并读取工作表名：" + ex);
                if (asSource) form.SetSourceSheets(path, null, ex.Message);
                else form.SetTargetSheets(null, ex.Message);
            }
            finally { _busy = false; }
        }

        // ---------------------------- 预检 ----------------------------

        private static void RunPreview(MergeHeadersForm form)
        {
            if (_busy) return;
            _busy = true;
            form.SetBusy(true);
            try
            {
                MergeRequest request = form.BuildRequest();
                string invalid = request.Validate();
                if (invalid != null) { form.SetLog(invalid); return; }

                if (_mergeGateway != null) { try { _mergeGateway.Dispose(); } catch { } }
                _mergeGateway = new MergeWorkbookGateway(_app);
                _mergeTarget = _mergeGateway.Resolve(request.Target.Path, request.Target.SheetName, false);
                _mergeTarget.Path = request.Target.Path;
                request.Target.SheetNames = MergeWorkbookGateway.SheetNames(_mergeTarget.Book);
                if (request.Target.SheetName.Length == 0) request.Target.SheetName = _mergeTarget.SheetName;
                if (!string.Equals(Path.GetExtension(request.Target.Path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("目标工作簿需为 .xlsx：" + request.Target.Path);

                form.SetStatus("正在读取目标表头…");
                MergeHeaderSet targetSet = MergeSheetReader.ReadHeaders(
                    _mergeTarget.Sheet, request.Target.HeaderRow, request.Target.Path);

                List<MergeHeaderSet> sourceSets = new List<MergeHeaderSet>();
                for (int i = 0; i < request.Sources.Count; i++)
                {
                    MergeWorkbookSpec spec = request.Sources[i];
                    form.SetStatus("正在读取来源 " + spec.DisplayName + "…");
                    MergeWorkbookHandle handle = _mergeGateway.Resolve(spec.Path, spec.SheetName, true);
                    spec.SheetNames = MergeWorkbookGateway.SheetNames(handle.Book);
                    if (spec.SheetName.Length == 0) spec.SheetName = handle.SheetName;
                    sourceSets.Add(MergeSheetReader.ReadHeaders(handle.Sheet, spec.HeaderRow, spec.Path));
                }

                MergeMatchResult match = MergeMatcher.Match(targetSet, sourceSets);
                if (match.Blocking.Count > 0)
                {
                    form.SetMatch(request, match);
                    return;
                }

                // 写入首行一律基于「写入前」的目标状态，因此先算完所有列再动笔。
                foreach (MergeSourcePlan plan in match.Sources)
                {
                    int common = -1;
                    foreach (MergeColumnMap map in plan.Maps)
                    {
                        int lastRow = MergeSheetReader.LastDataRow(_mergeTarget.Sheet, map.TargetColumn,
                            targetSet.HeaderRow, targetSet.LastUsedRow);
                        map.TargetStartRow = lastRow + 1;
                        if (common < 0) common = map.TargetStartRow;
                        else if (common != map.TargetStartRow) plan.StartRowsDiffer = true;
                    }
                }

                MergeReport previewReport = new MergeReport();
                Dictionary<string, RewriteResult> cache = new Dictionary<string, RewriteResult>(StringComparer.Ordinal);
                HashSet<string> issueKeys = new HashSet<string>(StringComparer.Ordinal);
                for (int i = 0; i < match.Sources.Count; i++)
                {
                    MergeSourcePlan plan = match.Sources[i];
                    MergeWorkbookSpec spec = request.Sources[i];
                    if (plan.Maps.Count == 0) continue;
                    form.SetStatus("正在读取 " + spec.DisplayName + " 的数据…");
                    MergeWorkbookHandle handle = _mergeGateway.Resolve(spec.Path, spec.SheetName, true);
                    MergeColumnMapper mapper = new MergeColumnMapper();
                    foreach (MergeColumnMap map in plan.Maps) mapper.Add(map.SourceColumn, map.TargetColumn);
                    HashSet<string> definedNames = ReadDefinedNames(handle.Book);
                    foreach (MergeColumnMap map in plan.Maps) MergeSheetReader.ReadColumn(handle.Sheet, map, spec.HeaderRow);
                    RewriteSource(plan, spec, mapper, definedNames, cache, issueKeys, previewReport);
                }

                string guard = MergePlanner.Guard(MergePlanner.Build(AllMaps(match), targetSet.HeaderRow));
                if (guard != null) { form.SetLog(guard); return; }

                form.SetMatch(request, match);
                form.SetSamples(previewReport.Samples);
                if (previewReport.Issues.Count > 0)
                    form.AppendLog("有 " + previewReport.Issues.Count + " 类引用未能完整重映射，已保持原样（详见执行后报告）。");
                _mergePreviewIssues = previewReport.Issues;
                Log("表头匹配合并预检：目标 " + request.Target.DisplayName + "，匹配 "
                    + match.MappedColumns + " 列，" + match.TotalCells + " 格");
            }
            catch (Exception ex)
            {
                Log("表头匹配合并预检：" + ex);
                form.SetLog("预检失败：" + ex.Message);
            }
            finally
            {
                form.SetBusy(false);
                _busy = false;
            }
        }

        private static List<string> _mergePreviewIssues = new List<string>();

        private static List<MergeColumnMap> AllMaps(MergeMatchResult match)
        {
            List<MergeColumnMap> all = new List<MergeColumnMap>();
            foreach (MergeSourcePlan plan in match.Sources) all.AddRange(plan.Maps);
            return all;
        }

        private static void RewriteSource(MergeSourcePlan plan, MergeWorkbookSpec spec,
            MergeColumnMapper mapper, HashSet<string> definedNames,
            Dictionary<string, RewriteResult> cache, HashSet<string> issueKeys, MergeReport report)
        {
            foreach (MergeColumnMap map in plan.Maps)
            {
                if (map.Cells == null) continue;
                RewriteContext context = new RewriteContext
                {
                    ColumnDelta = map.TargetColumn - map.SourceColumn,
                    SourcePath = spec.Path
                };
                for (int r = 0; r < map.RowCount; r++)
                {
                    object value = r < map.Cells.Length ? map.Cells[r] : null;
                    if (!MergePlanner.IsFormula(value)) continue;
                    string formula = Convert.ToString(value, CultureInfo.InvariantCulture);
                    string key = formula + "|" + map.RowDelta.ToString(CultureInfo.InvariantCulture)
                        + "|" + context.ColumnDelta.ToString(CultureInfo.InvariantCulture);
                    RewriteResult rewritten;
                    if (!cache.TryGetValue(key, out rewritten))
                    {
                        rewritten = FormulaRefRewriter.Rewrite(formula, mapper, map.RowDelta, definedNames, context);
                        cache.Add(key, rewritten);
                    }
                    map.Cells[r] = rewritten.Text;
                    if (rewritten.Changed) report.RewrittenFormulas++;
                    foreach (RefIssue issue in rewritten.Issues)
                        if (issueKeys.Add(issue.Key()))
                            report.Issues.Add(spec.DisplayName + "：" + issue.Describe());
                    if (rewritten.Changed && report.Samples.Count < 50)
                        report.Samples.Add(new MergeRewriteSample
                        {
                            Source = spec.DisplayName,
                            Cell = map.SourceLetter + (map.SourceFirstRow + r).ToString(CultureInfo.InvariantCulture),
                            Before = formula,
                            After = rewritten.Text,
                            Note = map.SourceLetter + "→" + map.TargetLetter
                                + "，行偏移 " + map.RowDelta.ToString(CultureInfo.InvariantCulture)
                        });
                }
            }
        }

        private static HashSet<string> ReadDefinedNames(object book)
        {
            HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (book == null) return names;
            using (ComBag bag = new ComBag())
            {
                try
                {
                    dynamic list = bag.Own(((dynamic)book).Names);
                    int count = Convert.ToInt32(list.Count);
                    for (int i = 1; i <= count; i++)
                        using (ComBag one = new ComBag())
                        {
                            dynamic item = one.Own(MergeWorkbookGateway.GetItem(list, i));
                            string text = Convert.ToString(item.Name);
                            if (string.IsNullOrEmpty(text)) continue;
                            int bang = text.IndexOf('!');
                            if (bang >= 0 && bang + 1 < text.Length) text = text.Substring(bang + 1);
                            if (text.Length > 0) names.Add(text);
                        }
                }
                catch { }
            }
            return names;
        }

        // ---------------------------- 执行 ----------------------------

        // 带 DisplayAlerts / AskToUpdateLinks 保护保存目标工作簿（避免带外部链接时弹更新/保存提示）。
        private static string SaveTargetPath()
        {
            dynamic a = _app;
            bool alerts = ComFormulaHost.Boolean(a.DisplayAlerts, "DisplayAlerts");
            bool ask = ComFormulaHost.Boolean(a.AskToUpdateLinks, "AskToUpdateLinks");
            try { a.DisplayAlerts = false; a.AskToUpdateLinks = false; return WorkbookSaving.Save(_mergeTarget.Book, true); }
            finally { a.AskToUpdateLinks = ask; a.DisplayAlerts = alerts; }
        }
        private static void SaveTarget()
        {
            dynamic a = _app;
            bool alerts = ComFormulaHost.Boolean(a.DisplayAlerts, "DisplayAlerts");
            bool ask = ComFormulaHost.Boolean(a.AskToUpdateLinks, "AskToUpdateLinks");
            try { a.DisplayAlerts = false; a.AskToUpdateLinks = false; WorkbookSaving.Save(_mergeTarget.Book, true); }
            finally { a.AskToUpdateLinks = ask; a.DisplayAlerts = alerts; }
        }

        /// <summary>
        /// 检测来源工作簿里所有外部链接（跨簿引用）指向的文件是否存在。
        /// 返回不存在的目标文件路径列表（去重、有序）。仅支持 .xlsx 来源。
        /// </summary>
        private static List<string> DetectMissingExternalLinks(MergeRequest request)
        {
            List<string> missing = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (MergeWorkbookSpec spec in request.Sources)
            {
                if (string.IsNullOrEmpty(spec.Path)) continue;
                List<ExternalLinkInfo> links;
                try { links = XlsxLinkPackage.Scan(spec.Path); }
                catch { continue; }   // 非 xlsx / 无法扫描的来源跳过（打开时已校验）
                foreach (ExternalLinkInfo link in links)
                {
                    if (string.IsNullOrEmpty(link.SourcePath)) continue;
                    if (!File.Exists(link.SourcePath) && seen.Add(link.SourcePath))
                        missing.Add(link.SourcePath);
                }
            }
            return missing;
        }

        private static void RunExecute(MergeHeadersForm form)
        {
            if (_busy) return;
            MergeMatchResult match = form.Match;
            if (match == null || match.Blocking.Count > 0 || match.MappedColumns == 0) return;
            _busy = true;
            form.SetBusy(true);
            MergeReport report = new MergeReport();
            try
            {
                MergeRequest request = form.BuildRequest();
                string invalid = request.Validate();
                if (invalid != null) { form.SetLog(invalid); return; }
                if (_mergeTarget == null || _mergeTarget.Sheet == null)
                    throw new InvalidOperationException("目标工作表已失效，请重新预检。");

                MergeWritePlan plan = MergePlanner.Build(AllMaps(match), request.Target.HeaderRow);
                string guard = MergePlanner.Guard(plan);
                if (guard != null) { form.SetLog(guard); return; }

                // 二次核对：表头未变、身份未变、映射仍在。
                MergeHeaderSet recheck = MergeSheetReader.ReadHeaders(
                    _mergeTarget.Sheet, request.Target.HeaderRow, request.Target.Path);
                string error = recheck.Validate("目标工作簿");
                if (error != null) throw new InvalidOperationException("目标表头已变化：" + error);

                // 迁移前检测来源工作簿引用的外部工作簿是否存在；缺失则询问用户决定。
                List<string> missingLinks = DetectMissingExternalLinks(request);
                if (missingLinks.Count > 0)
                {
                    string listing = string.Join("\r\n", missingLinks.ConvertAll(p => "· " + p));
                    Log("表头匹配合并：来源存在失效外部链接 " + missingLinks.Count + " 个");
                    if (UiDialog.Show(form,
                        "来源工作簿引用的下列外部工作簿不存在：\r\n\r\n" + listing
                        + "\r\n\r\n仍要继续迁移吗？（迁移会保留这些引用，目标工作簿打开时可能提示链接无法更新）",
                        "检测到失效的外部链接", System.Windows.Forms.MessageBoxButtons.OKCancel,
                        System.Windows.Forms.MessageBoxIcon.Warning,
                        System.Windows.Forms.MessageBoxDefaultButton.Button2) != System.Windows.Forms.DialogResult.OK)
                    {
                        form.AppendLog("已取消：来源存在失效外部链接。");
                        return;
                    }
                }

                Log("表头匹配合并：确认对话框前 host enabled=" + HostDialogService.HostEnabled());
                if (UiDialog.Show(form, MergePlanText.ConfirmationText(request, match),
                    "确认执行跨工作簿表头匹配合并",
                    System.Windows.Forms.MessageBoxButtons.OKCancel,
                    System.Windows.Forms.MessageBoxIcon.Warning,
                    System.Windows.Forms.MessageBoxDefaultButton.Button2) != System.Windows.Forms.DialogResult.OK)
                {
                    Log("表头匹配合并：确认对话框后 host enabled=" + HostDialogService.HostEnabled());
                    form.AppendLog("已取消，未写入任何内容。");
                    UiDialog.Show(form, "已取消，未写入任何内容。", "跨工作簿表头匹配合并",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }
                Log("表头匹配合并：确认通过，host enabled=" + HostDialogService.HostEnabled());

                form.SetStatus("正在备份目标工作簿…");
                string saved = SaveTargetPath();
                _mergeBackupPath = MergeBackup.Create(saved);

                form.SetStatus("正在快照写入区域…");
                _mergeSnapshot = MergeSnapshot.Capture(_mergeTarget.Sheet, plan);

                form.SetStatus("正在写入…");
                MergeTargetWriter.Write(_app, _mergeTarget.Sheet, plan, null);

                SaveTarget();

                report.Sources = match.Sources.Count;
                report.MappedColumns = match.MappedColumns;
                report.WrittenCells = plan.TotalCells;
                report.FormulaCells = plan.FormulaCells;
                foreach (string issue in _mergePreviewIssues) report.Issues.Add(issue);
                report.BackupPath = _mergeBackupPath;
                foreach (MergeSourcePlan source in match.Sources)
                {
                    report.SkippedColumns += source.Missing.Count;
                    foreach (MergeColumnMap map in source.Maps)
                        if (map.Cells != null)
                            for (int r = 0; r < map.RowCount; r++)
                                if (MergePlanner.IsFormula(map.Cells[r])) report.RewrittenFormulas++;
                }
                form.SetReport(report);
                form.SetStatus("写入完成；如需还原请点「取消操作」，确认无误请点「确认操作」。");
                Log("表头匹配合并写入：" + report.WrittenCells + " 格，备份 " + _mergeBackupPath);
            }
            catch (Exception ex)
            {
                Log("表头匹配合并执行：" + ex);
                report.BackupPath = _mergeBackupPath;
                form.SetReport(report);
                form.AppendLog("执行失败：" + ex.Message + "（目标工作表未改动或已按快照还原）");
                TryRollback(form);
                // 日志区不显眼；用模态框把真实失败原因直接端到用户面前，避免"以为执行了"。
                UiDialog.Show(form, "执行失败：" + ex.Message + "\r\n（目标工作表未改动或已按快照还原；详见日志。）",
                    "跨工作簿表头匹配合并", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            finally
            {
                form.SetBusy(false);
                _busy = false;
            }
        }

        private static void TryRollback(MergeHeadersForm form)
        {
            if (_mergeSnapshot == null || _mergeTarget == null || _mergeTarget.Sheet == null) return;
            try
            {
                MergeTargetWriter.Restore(_app, _mergeTarget.Sheet, _mergeSnapshot);
                SaveTarget();
                form.AppendLog("已用写入前快照还原目标工作表。");
            }
            catch (Exception ex)
            {
                Log("表头匹配合并还原失败：" + ex);
                form.AppendLog("还原失败：" + ex.Message + "；请用备份文件手动恢复：" + _mergeBackupPath);
            }
        }

        // ---------------------------- 取消 / 确认 ----------------------------

        private static void RunUndo(MergeHeadersForm form)
        {
            if (_busy) return;
            _busy = true;
            form.SetBusy(true);
            try
            {
                if (_mergeSnapshot == null || _mergeTarget == null || _mergeTarget.Sheet == null)
                {
                    form.AppendLog("没有可还原的批次。");
                    return;
                }
                if (UiDialog.Show(form, "将用写入前的快照还原目标工作表，并自动保存。是否继续？",
                    "取消操作", System.Windows.Forms.MessageBoxButtons.OKCancel,
                    System.Windows.Forms.MessageBoxIcon.Warning,
                    System.Windows.Forms.MessageBoxDefaultButton.Button2) != System.Windows.Forms.DialogResult.OK)
                {
                    form.AppendLog("已取消还原。");
                    return;
                }
                MergeTargetWriter.Restore(_app, _mergeTarget.Sheet, _mergeSnapshot);
                SaveTarget();
                _mergeSnapshot = null;
                form.MarkUndone();
                form.AppendLog("已还原到写入前状态。文件备份保留：" + _mergeBackupPath);
                Log("表头匹配合并：已还原");
            }
            catch (Exception ex)
            {
                Log("表头匹配合并取消操作：" + ex);
                form.AppendLog("还原失败：" + ex.Message + "；请用备份文件手动恢复：" + _mergeBackupPath);
            }
            finally
            {
                form.SetBusy(false);
                _busy = false;
            }
        }

        private static void RunConfirm(MergeHeadersForm form)
        {
            _mergeSnapshot = null;
            Log("表头匹配合并：已确认，本批结束。备份保留于 " + _mergeBackupPath);
        }
    }
}
