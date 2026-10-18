using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // =====================================================================
    // 功能区与业务调度（VSTO 版，挂在 WPS 表格 ET）
    //
    //  - 实现 Microsoft.Office.Core.IRibbonExtensibility
    //    （IID 000C0396-...，与 WPS/Office 一致，见 InteropDefs.cs）
    //  - ThisAddIn.CreateRibbonExtensibilityObject() 返回本类实例，
    //    vstor 经其 GetCustomUI 取功能区 XML，onAction / getImage 经
    //    IDispatch 晚期绑定回调（AutoDispatch 保证 GetIDsOfNames 可解析）
    //  - WPS 表格 Application 经 Marshal.GetActiveObject 懒获取并探活，
    //    OnShutdown 时由 ThisAddIn 置空
    //
    // 按钮专属回调名原则（参考项目实测）：函数名即身份，不读 control.Id
    // （网络盘文档回调代理读不出 Id，靠 id 分发会把所有按钮认成同一个）。
    // 新增按钮 = Configure_program.py 加条目 + 这里加一行包装（配置器校验存在性）。
    // =====================================================================
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public partial class Connect : Microsoft.Office.Core.IRibbonExtensibility
    {
        private static object _app;
        private static bool _busy;
        private static Timer _formulaStatusTimer;
        private static object _formulaStatusApp, _formulaStatusPrevious;
        private static string _formulaStatusText;
        private static TrackerForm _trackerForm;   // 「增强追踪」非模态窗（全程只开一个）
        private static ColumnFilterForm _columnForm;
        private static OrganizeFilesForm _organizeForm;
        private static RenameFilesForm _renameForm;
        private static LinkAssistantForm _linkForm;
        private static DirectoryNavigation _directoryNavigation;
        private static Timer _hostConnectionTimer;
        internal static void StartHostConnection()
        {
            StopHostConnection(); int attempts = 0;
            _hostConnectionTimer = new Timer { Interval = 1000 };
            _hostConnectionTimer.Tick += delegate {
                attempts++;
                if (EnsureApplication() && _directoryNavigation == null) try { _directoryNavigation = new DirectoryNavigation(_app); } catch (Exception ex) { Log("目录事件重试：" + ex.Message); }
                if (_directoryNavigation != null || attempts >= 30) StopHostConnection();
            };
            _hostConnectionTimer.Start();
        }
        internal static void StopHostConnection()
        { if (_hostConnectionTimer != null) { _hostConnectionTimer.Stop(); _hostConnectionTimer.Dispose(); _hostConnectionTimer = null; } }

        // ----------------------- WPS 表格 Application -----------------------

        /// <summary>VSTO 入口在 OnShutdown 时置空；按钮点击时也可惰性重建。</summary>
        public static object HostApplication
        {
            get { return _app; }
            set
            {
                if (!ReferenceEquals(_app, value))
                {
                    FormulaInteraction.Shutdown();
                    RestoreFormulaStatus();
                    if (_directoryNavigation != null) { try { _directoryNavigation.Dispose(); } catch (Exception ex) { Log("目录事件解绑：" + ex.Message); } _directoryNavigation = null; }
                    if (_linkForm != null && !_linkForm.IsDisposed) _linkForm.Close();
                    if (_columnForm != null && !_columnForm.IsDisposed) _columnForm.Close();
                    if (_organizeForm != null && !_organizeForm.IsDisposed) _organizeForm.Close();
                    if (_renameForm != null && !_renameForm.IsDisposed) _renameForm.Close();
                    if (_mergeForm != null && !_mergeForm.IsDisposed) _mergeForm.Close();
                    if (_eliminationForm != null && !_eliminationForm.IsDisposed) _eliminationForm.Close();
                }
                _app = value;
                if (_app != null && _directoryNavigation == null) try { _directoryNavigation = new DirectoryNavigation(_app); } catch (Exception ex) { Log("目录事件连接：" + ex.Message); }
            }
        }

        /// <summary>确保拿到运行中的 WPS 表格 Application；拿不到返回 false。</summary>
        public static bool EnsureApplication()
        {
            if (IsAlive(_app)) return true;
            HostApplication = null;
            // ProgID 回退链：WPS 表格首选项 KET.Application（开发日志实证），
            // 其次 et.Application / Excel.Application。
            string[] progIds = { "KET.Application", "et.Application", "Excel.Application" };
            for (int i = 0; i < progIds.Length; i++)
            {
                try
                {
                    object o = Marshal.GetActiveObject(progIds[i]);
                    if (o != null)
                    {
                        HostApplication = o;
                        Log("EnsureApplication: 经 " + progIds[i] + " 获取成功");
                        return true;
                    }
                }
                catch { }
            }
            Log("EnsureApplication: 未找到运行中的 WPS 表格实例");
            return false;
        }

        private static bool IsAlive(object app)
        {
            if (app == null) return false;
            object workbooks = null;
            try
            {
                // ET 用 Workbooks 探活（无 Documents 集合，与 Word 不同）
                dynamic d = app;
                workbooks = d.Workbooks;
                return workbooks != null;
            }
            catch { return false; }
            finally { ReleaseTemporary(workbooks); }
        }

        private static void ReleaseTemporary(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return;
            try { Marshal.ReleaseComObject(value); }
            catch (Exception ex) { Log("临时 COM 引用释放失败：" + ex.Message); }
        }

        // ------------------------ IRibbonExtensibility ------------------------

        // Ribbon XML 由 Configure_program.py 依据「按钮配置区」生成，
        // 输出到 src/addin/Ribbon.g.cs —— 改标题/提示/顺序请改配置后重跑脚本。
        public string GetCustomUI(string RibbonID)
        {
            Log("GetCustomUI ribbonId=" + (RibbonID ?? ""));
            return RibbonCustomization.Effective(true).OuterXml;
        }

        // --------------------------- Ribbon 回调 ---------------------------
        // 每按钮专属 getImage 回调（函数名即身份，不读 control.Id）。
        // 蓝色（追踪引用）
        public IPictureDisp GetImage_btnTracePrecedents(object control)
        { return GetTraceImage("btnTracePrecedents", false); }

        // 绿色（追踪从属）
        public IPictureDisp GetImage_btnTraceDependents(object control)
        { return GetTraceImage("btnTraceDependents", true); }

        public IPictureDisp GetImage_btnFeedback(object control)
        { return GetButtonImageCore("btnFeedback"); }

        public IPictureDisp GetImage_menuOneClickFormula(object control) { return GetButtonImageCore("menuOneClickFormula"); }
        public IPictureDisp GetImage_menuWrapFormula(object control) { return GetButtonImageCore("menuWrapFormula"); }
        public IPictureDisp GetImage_menuIFErrorWrap(object control) { return GetButtonImageCore("menuIFErrorWrap"); }
        public IPictureDisp GetImage_menuRoundWrap(object control) { return GetButtonImageCore("menuRoundWrap"); }
        public IPictureDisp GetImage_btnWrapIFError(object control) { return GetButtonImageCore("btnWrapIFError"); }
        public IPictureDisp GetImage_btnWrapAbs(object control) { return GetButtonImageCore("btnWrapAbs"); }
        public IPictureDisp GetImage_btnWrapRound(object control) { return GetButtonImageCore("btnWrapRound"); }
        public IPictureDisp GetImage_btnCustomIFError(object control) { return GetButtonImageCore("btnCustomIFError"); }
        public IPictureDisp GetImage_btnCustomRound(object control) { return GetButtonImageCore("btnCustomRound"); }
        public IPictureDisp GetImage_btnUnwrapFormula(object control) { return GetButtonImageCore("btnUnwrapFormula"); }
        public IPictureDisp GetImage_btnCustomFunction(object control) { return GetButtonImageCore("btnCustomFunction"); }

        /// <summary>统一的图标分发：优先用配置里的自定义图片，没有则按配置回落内置图标。</summary>
        private static IPictureDisp GetTraceImage(string id, bool green)
        {
            IPictureDisp custom = RibbonConfig.GetImage(id);
            if (custom != null) return custom;
            return MakeRibbonIcon(green ? "dependents" : "precedents", green ? Ui.Green : Ui.Blue, RibbonConfig.GetImageSize(id));
        }

        private static IPictureDisp GetButtonImageCore(string id)
        {
            Color accent = id == "btnFeedback" ? Ui.Amber : Ui.Blue;
            try
            {
                if (string.IsNullOrEmpty(id)) return MakeRibbonIcon("precedents", Ui.Blue, 32);
                IPictureDisp custom = RibbonConfig.GetImage(id);
                if (custom != null) return custom;
                string kind = RibbonConfig.BuiltinIconKind(id);
                if (kind == RibbonConfig.BuiltinView) return MakeRibbonIcon("view", accent, RibbonConfig.GetImageSize(id));
                if (kind == RibbonConfig.BuiltinFilter) return MakeRibbonIcon("filter", accent, RibbonConfig.GetImageSize(id));
                if (kind == RibbonConfig.BuiltinFormula) return MakeRibbonIcon("formula", accent, RibbonConfig.GetImageSize(id));
                if (kind == RibbonConfig.BuiltinLink || kind == RibbonConfig.BuiltinDirectory || kind == RibbonConfig.BuiltinAgent || kind == RibbonConfig.BuiltinWeb || kind == RibbonConfig.BuiltinDocument || kind == RibbonConfig.BuiltinSettings) return MakeRibbonIcon(kind, accent, RibbonConfig.GetImageSize(id));
                if (kind == RibbonConfig.BuiltinMail) return MakeRibbonIcon("mail", accent, RibbonConfig.GetImageSize(id));
                if (kind == RibbonConfig.BuiltinUpdate || kind == RibbonConfig.BuiltinPaste || kind == RibbonConfig.BuiltinTable)
                    return MakeRibbonIcon(kind, accent, RibbonConfig.GetImageSize(id));
                return MakeRibbonIcon("precedents", accent, RibbonConfig.GetImageSize(id));
            }
            catch (Exception ex)
            {
                Log("GetButtonImageCore(" + (id ?? "") + ") 异常：" + ex.Message);
                return null;
            }
        }

        public void OnCalculateAdd(object control) { RunFormula(FormulaOperation.Add); }
        public void OnCustomCalculateAdd(object control) { RunFormula(FormulaOperation.Add, true); }
        public void OnCalculateSubtract(object control) { RunFormula(FormulaOperation.Subtract); }
        public void OnCustomCalculateSubtract(object control) { RunFormula(FormulaOperation.Subtract, true); }
        public void OnCalculateMultiply(object control) { RunFormula(FormulaOperation.Multiply); }
        public void OnCustomCalculateMultiply(object control) { RunFormula(FormulaOperation.Multiply, true); }
        public void OnCalculateDivide(object control) { RunFormula(FormulaOperation.Divide); }
        public void OnCustomCalculateDivide(object control) { RunFormula(FormulaOperation.Divide, true); }
        public IPictureDisp GetImage_menuFilterEnhanced(object control) { return GetButtonImageCore("menuFilterEnhanced"); }
        public IPictureDisp GetImage_btnFilterColumns(object control) { return GetButtonImageCore("btnFilterColumns"); }
        public IPictureDisp GetImage_menuArithmetic(object control) { return GetButtonImageCore("menuArithmetic"); }
        public IPictureDisp GetImage_menuCalculateAdd(object control) { return GetButtonImageCore("menuCalculateAdd"); }
        public IPictureDisp GetImage_btnCalculateAdd(object control) { return GetButtonImageCore("btnCalculateAdd"); }
        public IPictureDisp GetImage_btnCustomCalculateAdd(object control) { return GetButtonImageCore("btnCustomCalculateAdd"); }
        public IPictureDisp GetImage_menuCalculateSubtract(object control) { return GetButtonImageCore("menuCalculateSubtract"); }
        public IPictureDisp GetImage_btnCalculateSubtract(object control) { return GetButtonImageCore("btnCalculateSubtract"); }
        public IPictureDisp GetImage_btnCustomCalculateSubtract(object control) { return GetButtonImageCore("btnCustomCalculateSubtract"); }
        public IPictureDisp GetImage_menuCalculateMultiply(object control) { return GetButtonImageCore("menuCalculateMultiply"); }
        public IPictureDisp GetImage_btnCalculateMultiply(object control) { return GetButtonImageCore("btnCalculateMultiply"); }
        public IPictureDisp GetImage_btnCustomCalculateMultiply(object control) { return GetButtonImageCore("btnCustomCalculateMultiply"); }
        public IPictureDisp GetImage_menuCalculateDivide(object control) { return GetButtonImageCore("menuCalculateDivide"); }
        public IPictureDisp GetImage_btnCalculateDivide(object control) { return GetButtonImageCore("btnCalculateDivide"); }
        public IPictureDisp GetImage_btnCustomCalculateDivide(object control) { return GetButtonImageCore("btnCustomCalculateDivide"); }

        public void OnWrapIFError(object control) { RunFormula(FormulaOperation.IFError); }
        public void OnWrapAbs(object control) { RunFormula(FormulaOperation.Abs); }
        public void OnWrapRound(object control) { RunFormula(FormulaOperation.Round); }
        public void OnUnwrapFormula(object control) { RunFormula(FormulaOperation.Unwrap); }

        public void OnCustomIFError(object control) { RunFormula(FormulaOperation.IFError, true); }
        public void OnCustomRound(object control) { RunFormula(FormulaOperation.Round, true); }
        public void OnCustomFunction(object control) { RunFormula(FormulaOperation.Custom, true); }

        public IPictureDisp GetImage_btnQuickAdvancedFilter(object control) { return GetButtonImageCore("btnQuickAdvancedFilter"); }
        public IPictureDisp GetImage_btnLinkAssistant(object control) { return GetButtonImageCore("btnLinkAssistant"); }
        public IPictureDisp GetImage_btnGenerateDirectory(object control) { return GetButtonImageCore("btnGenerateDirectory"); }
        internal static void CloseWorkbookWindows()
        { if (_columnForm != null && !_columnForm.IsDisposed) _columnForm.Close(); if (_trackerForm != null && !_trackerForm.IsDisposed) _trackerForm.Close(); }
        internal static void PrepareWorkbookSave(string path)
        { if (_trackerForm != null && !_trackerForm.IsDisposed) _trackerForm.CleanupForWorkbook(path); }
        public IPictureDisp GetImage_menuFileOperations(object control) { return GetButtonImageCore("menuFileOperations"); }
        public IPictureDisp GetImage_btnOrganizeFiles(object control) { return GetButtonImageCore("btnOrganizeFiles"); }
        public IPictureDisp GetImage_btnRenameFiles(object control) { return GetButtonImageCore("btnRenameFiles"); }
        private static void ShowFileForm(Form form)
        { IntPtr owner = GetHostWindowHandle(); if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner)); }
        public void OnOrganizeFiles(object control)
        {
            if (_busy) return;
            if (_organizeForm != null && !_organizeForm.IsDisposed) { _organizeForm.Activate(); return; }
            try { var form = new OrganizeFilesForm(); _organizeForm = form; form.FormClosed += delegate { _organizeForm = null; form.Dispose(); }; ShowFileForm(form); }
            catch (Exception ex) { Log("文件批量整理："+ex); UiDialog.Show(ex.Message,"文件批量整理",MessageBoxButtons.OK,MessageBoxIcon.Warning); }
        }
        public void OnRenameFiles(object control)
        {
            if (_busy) return;
            if (_renameForm != null && !_renameForm.IsDisposed) { _renameForm.Activate(); return; }
            try { if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。"); var form = new RenameFilesForm(new RenameWorkbookSession(_app)); _renameForm = form; form.FormClosed += delegate { _renameForm = null; form.Dispose(); }; ShowFileForm(form); }
            catch (Exception ex) { Log("文件批量命名："+ex); UiDialog.Show(ex.Message,"文件批量命名",MessageBoxButtons.OK,MessageBoxIcon.Warning); }
        }
        public void OnQuickAdvancedFilter(object control)
        {
            if (_busy) return; _busy = true;
            try { if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。"); int count = QuickAdvancedFilter.Run(_app); ShowStatus("快捷高筛：已输出 " + count + " 条记录至新工作表"); }
            catch (Exception ex) { Log("快捷高筛：" + ex); UiDialog.Show(ex.Message, "快捷高筛", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { _busy = false; }
        }
        public void OnGenerateDirectory(object control)
        {
            if (_busy) return; _busy = true;
            try { if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。"); if (_directoryNavigation == null) _directoryNavigation = new DirectoryNavigation(_app); int count = DirectoryNavigation.Generate(_app); ShowStatus("生成目录：已链接 " + count + " 张工作表"); }
            catch (Exception ex) { Log("生成目录：" + ex); UiDialog.Show(ex.Message, "生成目录", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { _busy = false; }
        }
        public void OnLinkAssistant(object control)
        {
            if (_busy) return;
            if (_linkForm != null && !_linkForm.IsDisposed) { _linkForm.Activate(); return; }
            try { if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。"); LinkAssistantForm form = new LinkAssistantForm(new LinkAssistantSession(_app)); _linkForm = form; form.FormClosed += delegate { _linkForm = null; form.Dispose(); }; IntPtr owner = GetHostWindowHandle(); if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner)); }
            catch (Exception ex) { Log("链接助手：" + ex); UiDialog.Show(ex.Message, "链接助手", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        public void OnFilterColumns(object control)
        {
            if (_busy) return;
            if (_columnForm != null && !_columnForm.IsDisposed) { _columnForm.Activate(); return; }
            _busy = true;
            ComColumnVisibilityHost host = null;
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格并选中一行。");
                host = new ComColumnVisibilityHost(_app);
                var columns = host.Collect();
                ColumnFilterForm form = new ColumnFilterForm(columns, host.RangeDescription);
                ComColumnVisibilityHost bound = host;
                form.ApplyRequested += delegate(System.Collections.Generic.IList<ColumnChoice> choices)
                {
                    if (_busy) return;
                    _busy = true;
                    try
                    {
                        FormulaCommandResult result = ColumnVisibilityEngine.Apply(bound, choices,
                            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Excel_Iterate", "ColumnRecovery"));
                        if (!result.Failed)
                        {
                            if (result.Changed > 0) ShowStatus("按列筛选：已修改 " + result.Changed + " 列的显示状态");
                            form.Close();
                        }
                        else { Log(result.Details); UiDialog.Show(form, result.Details, "按列筛选", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                    }
                    finally { _busy = false; }
                };
                form.FormClosed += delegate { _columnForm = null; try { bound.Dispose(); } catch (Exception ex) { Log("列筛选引用释放：" + ex.Message); } form.Dispose(); };
                _columnForm = form;
                IntPtr owner = GetHostWindowHandle();
                if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
                host = null; // The non-modal form owns the snapshot until it closes.
            }
            catch (Exception ex) { Log("按列筛选：" + ex); UiDialog.Show(ex.Message, "按列筛选", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { if (host != null) host.Dispose(); _busy = false; }
        }

        private static FormulaCapture CaptureFor(FormulaOperation operation)
        { return operation == FormulaOperation.Custom ? FormulaCapture.Values : operation >= FormulaOperation.Add ? FormulaCapture.Numbers : FormulaCapture.Formulas; }
        private static void RunFormula(FormulaOperation operation) { RunFormula(operation, false); }
        internal static bool TryBeginFormulaExecution() { if (_busy) return false; _busy = true; return true; }
        internal static void EndFormulaExecution() { _busy = false; }
        private static void RunFormula(FormulaOperation operation, bool custom)
        {
            if (custom && FormulaInteraction.ActivatePending()) return;
            if (!TryBeginFormulaExecution()) return;
            IntPtr owner = IntPtr.Zero;
            try
            {
                if (!EnsureApplication())
                {
                    FormulaWindowThread.Notice("请先打开 WPS 表格工作簿并选中单元格。", owner, true);
                    return;
                }
                owner = FormulaWindowState.ResolveOwner(_app);
                if (custom)
                {
                    // Open takes ownership. Waiting for parameters never holds _busy
                    // and never starts a modal message loop on the Ribbon thread.
                    FormulaInteraction.Open(new ComFormulaHost(_app, CaptureFor(operation)), operation, owner, _app);
                    return;
                }
                using (ComFormulaHost host = new ComFormulaHost(_app, CaptureFor(operation)))
                {
                    PresentFormulaResult(FormulaCommandEngine.Run(host, FormulaRequest.Default(operation), Log,
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Excel_Iterate", "FormulaRecovery")), owner);
                }
            }
            catch (Exception ex)
            {
                UiDiagnostics.Record("formula-command", ex);
                FormulaWindowThread.Notice("一键公式发生错误：" + ex.Message, owner, true);
            }
            finally { EndFormulaExecution(); }
        }
        internal static void PresentFormulaResult(FormulaCommandResult result, IntPtr owner)
        {
            if (!result.Failed && result.Changed > 0) ShowFormulaStatus(result.Changed);
            else FormulaWindowThread.Notice(result.Details, owner, result.Failed);
        }
        private static void ShowFormulaStatus(int count)
        { ShowStatus("一键公式：已修改 " + count + " 个单元格"); }
        private static void ShowStatus(string message)
        {
            RestoreFormulaStatus();
            try
            {
                _formulaStatusApp = _app;
                _formulaStatusPrevious = ((dynamic)_app).StatusBar;
                _formulaStatusText = message;
                ((dynamic)_app).StatusBar = _formulaStatusText;
                _formulaStatusTimer = new Timer { Interval = 5000 };
                _formulaStatusTimer.Tick += delegate { RestoreFormulaStatus(); };
                _formulaStatusTimer.Start();
            }
            catch (Exception ex) { Log(message + "；状态栏提示失败：" + ex.Message); RestoreFormulaStatus(); }
        }
        private static void RestoreFormulaStatus()
        {
            if (_formulaStatusTimer != null) { _formulaStatusTimer.Stop(); _formulaStatusTimer.Dispose(); _formulaStatusTimer = null; }
            try
            {
                // Do not overwrite a newer status message from the host or another add-in.
                if (_formulaStatusApp != null && string.Equals(Convert.ToString(((dynamic)_formulaStatusApp).StatusBar), _formulaStatusText, StringComparison.Ordinal))
                    ((dynamic)_formulaStatusApp).StatusBar = _formulaStatusPrevious;
            }
            catch (Exception ex) { Log("恢复状态栏失败：" + ex.Message); }
            finally { _formulaStatusApp = _formulaStatusPrevious = null; _formulaStatusText = null; }
        }

        // --------------------------- 追踪回调 ---------------------------

        public void OnTracePrecedents(object control)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                if (!EnsureApplication())
                {
                    UiDialog.Show(
                        "未检测到正在运行的 WPS 表格。\r\n请先打开 WPS 表格工作簿，再点击「追踪引用」。",
                        "追踪引用", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                OpenTracker(TrackerMode.Precedents);
            }
            catch (Exception ex)
            {
                Log("OnTracePrecedents 异常：" + ex);
                try
                {
                    UiDialog.Show("发生未预期的错误：" + ex.Message, "追踪引用",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
            }
            finally { _busy = false; }
        }

        public void OnTraceDependents(object control)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                if (!EnsureApplication())
                {
                    UiDialog.Show(
                        "未检测到正在运行的 WPS 表格。\r\n请先打开 WPS 表格工作簿，再点击「追踪从属」。",
                        "追踪从属", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                OpenTracker(TrackerMode.Dependents);
            }
            catch (Exception ex)
            {
                Log("OnTraceDependents 异常：" + ex);
                try
                {
                    UiDialog.Show("发生未预期的错误：" + ex.Message, "追踪从属",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
            }
            finally { _busy = false; }
        }

        // ------------------ 网址回调（每个 url 按钮一个专属入口） ------------------
        //
        // 与图标同理：onAction="OnOpenUrl_<按钮id>"，回调名带来按钮身份。
        // 将来新增网址按钮时同样只加一行包装，不需要任何"唯一网址兜底"。

        public void OnOpenUrl_btnFeedback(object control) { OpenUrlCore("btnFeedback"); }

        /// <summary>统一的打开网页/邮箱分发（行为与地址在配置区定义）。</summary>
        private static void OpenUrlCore(string id)
        {
            try
            {
                string url = RibbonConfig.GetUrl(id);
                if (string.IsNullOrEmpty(url))
                {
                    Log("OnOpenUrl：按钮 " + (id ?? "") + " 未配置网址");
                    return;
                }
                Log("OnOpenUrl -> " + id + " " + url);
                System.Diagnostics.Process.Start(url);
            }
            catch (Exception ex)
            {
                Log("OnOpenUrl 异常：" + ex.Message);
                try
                {
                    UiDialog.Show("无法打开配置的网址。\r\n\r\n错误信息：" + ex.Message,
                        "Excel-iterate", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch { }
            }
        }

        /// <summary>打开追踪窗口：同模式已开则带到前台，不同模式先关旧开新。</summary>
        private static void OpenTracker(TrackerMode mode)
        {
            // 已开着同模式的窗 → 带到前台
            if (_trackerForm != null && !_trackerForm.IsDisposed && _trackerForm.Mode == mode)
            {
                try { _trackerForm.Activate(); } catch { }
                return;
            }
            // 关旧窗（若有，且模式不同）
            TrackerForm old = _trackerForm;
            _trackerForm = null;
            if (old != null && !old.IsDisposed)
            {
                try
                {
                    if (old.IsHandleCreated)
                        old.BeginInvoke((MethodInvoker)delegate { try { old.Close(); } catch { } });
                }
                catch { try { old.Close(); } catch { } }
            }
            // 开新窗
            Log("OpenTracker：打开 " + mode + " 窗口");
            TrackerForm form = new TrackerForm(_app, mode);
            form.FormClosed += delegate { if (_trackerForm == form) _trackerForm = null; };
            _trackerForm = form;
            // 与链接助手/按列筛选/表头匹配合并完全一致：取宿主主窗口作 owner 显示；
            // owner 取不到就普通 Show()，绝不设 TopMost——无主 TopMost 窗会持续
            // 抢占 WPS 前台，表现为单元格无法输入、功能区子菜单点不开。
            IntPtr owner = GetHostWindowHandle();
            Log("增强追踪：打开前 enabled=" + HostDialogService.HostEnabled()
                + " / 前台=" + NativeOle.GetForegroundWindow() + " / owner=" + owner);
            if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
            Log("增强追踪：Show 后 enabled=" + HostDialogService.HostEnabled()
                + " / 前台=" + NativeOle.GetForegroundWindow()
                + " / 弹窗=" + (form.IsHandleCreated ? form.Handle : IntPtr.Zero));
        }

        /// <summary>卸载/关文档时关掉还开着的追踪窗，避免挂在已死的宿主上。</summary>
        public static void CloseTrackerForm()
        {
            TrackerForm f = _trackerForm;
            _trackerForm = null;
            if (f == null) return;
            try
            {
                if (!f.IsDisposed && f.IsHandleCreated)
                    f.BeginInvoke((MethodInvoker)delegate { try { f.Close(); } catch { } });
            }
            catch
            {
                try { f.Close(); } catch { }
            }
        }

        /// <summary>尽力拿 WPS 表格主窗口句柄，给非模态窗当 Owner 用（四级兜底）。</summary>
        private static IntPtr GetHostWindowHandle()
        {
            // 首选：宿主 Application 自己的主窗口句柄（Excel/WPS 均有 Hwnd），
            // 比 ActiveWindow 更稳——没有活动簿时 ActiveWindow 会抛。
            try
            {
                IntPtr handle = HostDialogService.ValidateOwner(new IntPtr(Convert.ToInt64(((dynamic)_app).Hwnd)));
                if (handle != IntPtr.Zero) return handle;
            }
            catch { }
            object window = null;
            try
            {
                window = ((dynamic)_app).ActiveWindow;
                IntPtr handle = HostDialogService.ValidateOwner(new IntPtr(Convert.ToInt64(((dynamic)window).Hwnd)));
                if (handle != IntPtr.Zero) return handle;
            }
            catch { }
            finally { ReleaseTemporary(window); }
            try
            {
                IntPtr handle = HostDialogService.ValidateOwner(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
                if (handle != IntPtr.Zero) return handle;
                return HostDialogService.ValidateOwner(NativeOle.GetForegroundWindow());
            }
            catch { return IntPtr.Zero; }
        }

        /// <summary>把 Win32 句柄包成 WinForms 认的 IWin32Window。</summary>
        private sealed class WindowWrapper : System.Windows.Forms.IWin32Window
        {
            private readonly IntPtr _h;
            public WindowWrapper(IntPtr h) { _h = h; }
            public IntPtr Handle { get { return _h; } }
        }

        // --------------------------- 按钮图标 ---------------------------

        private static IPictureDisp MakeRibbonIcon(string kind, Color accent, int logicalSize)
        {
            return RibbonImages.Create(logicalSize,
                delegate(int width, int height) { return UiIcons.CreateRibbon(kind, accent, width, height); }, Log);
        }

        // --------------------------- 日志 ---------------------------

        private static readonly object _logLock = new object();

        /// <summary>
        /// 固定日志位置 %LocalAppData%\Excel_Iterate\Excel_Iterate.log。
        /// VSTO 加载时 Assembly.Location 指向 CLR 的 dl3 影子缓存而非部署目录，
        /// 因此写到 %LocalAppData% 下。超 512KB 自动清空。
        /// </summary>
        public static void Log(string msg)
        {
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Excel_Iterate");
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "Excel_Iterate.log");
                lock (_logLock)
                {
                    FileInfo fi = new FileInfo(path);
                    if (fi.Exists && fi.Length > 512 * 1024)
                        File.WriteAllText(path, "");
                    File.AppendAllText(path,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + msg + "\r\n");
                }
            }
            catch { }
        }
    }
}
