using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace WordUpdater
{
    // =====================================================================
    // 功能区与业务调度（VSTO 版）
    //
    //  - 实现 Microsoft.Office.Core.IRibbonExtensibility
    //    （该接口内嵌于 Microsoft.Office.Tools.Common.v4.0.Utilities.dll，
    //      IID 000C0396-0000-0000-C000-000000000046，与 WPS/Office 一致）
    //  - ThisAddIn.CreateRibbonExtensibilityObject() 返回本类实例，
    //    vstor 经其 GetCustomUI 取功能区 XML，onAction / getImage 经
    //    IDispatch 晚期绑定回调（AutoDispatch 保证 GetIDsOfNames 可解析）
    //  - WPS Application 经 Marshal.GetActiveObject 懒获取并探活，
    //    OnShutdown 时由 ThisAddIn 置空
    // =====================================================================
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public partial class Connect : Microsoft.Office.Core.IRibbonExtensibility
    {
        private static object _app;
        private static bool _busy;
        private static LinkViewForm _viewForm;   // 「查看链接」窗口（非模态，全程只开一个）

        // 标黄“闪一下”的多槽容器（「更新链接」批量 + 「查看链接」单条共用）。
        // ★ 两套机制必须共用同一个 session：否则各自把自己的 Timer 到点时
        //   把【对方刚打上的黄】读成"原值"，最后还原成 wdYellow，黄色就永久
        //   留在用户文档里了。共用后由 Add() 按引用去重、保留最早的原值。
        private static FlashSession _flash;

        public static FlashSession SharedFlash
        {
            get
            {
                if (_flash == null) _flash = new FlashSession();
                return _flash;
            }
        }

        /// <summary>还原还没到期的标黄（新一轮开始前 / 加载项卸载时调用）。</summary>
        public static void ClearPendingFlash()
        {
            try
            {
                if (_flash != null)
                {
                    _flash.ClearNow();
                    _flash = null;
                }
            }
            catch { }
        }

        // ----------------------- WPS Application -----------------------

        /// <summary>VSTO 入口在 OnShutdown 时置空；按钮点击时也可惰性重建。</summary>
        public static object HostApplication
        {
            get { return _app; }
            set { _app = value; }
        }

        /// <summary>确保拿到运行中的 WPS 文字 Application；拿不到返回 false。</summary>
        public static bool EnsureApplication()
        {
            if (IsAlive(_app)) return true;
            _app = null;
            string[] progIds = { "WPS.Application", "KWPS.Application", "Word.Application" };
            for (int i = 0; i < progIds.Length; i++)
            {
                try
                {
                    object o = Marshal.GetActiveObject(progIds[i]);
                    if (o != null)
                    {
                        _app = o;
                        Log("EnsureApplication: 经 " + progIds[i] + " 获取成功");
                        return true;
                    }
                }
                catch { }
            }
            Log("EnsureApplication: 未找到运行中的 WPS 实例");
            return false;
        }

        private static bool IsAlive(object app)
        {
            if (app == null) return false;
            object documents = null;
            try
            {
                // 轻量属性访问探活（COM 对象已分离会抛异常）
                dynamic d = app;
                documents = d.Documents;
                return documents != null;
            }
            catch { return false; }
            finally { ReleaseTemporary(documents); }
        }

        private static void ReleaseTemporary(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return;
            try { Marshal.ReleaseComObject(value); }
            catch (Exception ex) { Log("临时 COM 引用释放失败：" + ex.Message); }
        }

        // ------------------------ IRibbonExtensibility ------------------------

        // Ribbon XML 由 配置器脚本 依据脚本顶部的「按钮配置区」生成，
        // 输出到 src/addin/Ribbon.g.cs —— 改标题/提示/顺序请改配置后重跑脚本，
        // 不要手改本文件（回归：本类只负责把回调转发给具体功能）。
        public string GetCustomUI(string RibbonID)
        {
            Log("GetCustomUI ribbonId=" + (RibbonID ?? ""));
            return RibbonCustomization.Effective(true).OuterXml;
        }

        // --------------------------- Ribbon 回调 ---------------------------

        // ------------------ 图标回调（每按钮一个专属入口） ------------------
        //
        // 每个按钮在 Ribbon XML 里都有【专属 getImage 回调名】（GetImage_<按钮id>），
        // 函数名即按钮身份 —— 不再读 control.Id（2026-09-26 日志实测：网络盘
        // 文档的回调代理按会话间歇性读不出 Id，靠 id 分发会把所有按钮认成同一个，
        // 图标全白/全一样）。本地与网络盘文档下表现一致。
        //
        // 新增按钮 = 配置器脚本 加条目 + 这里加一行包装（配置器会校验，漏加当场报错）。

        public IPictureDisp GetImage_btnUpdateLinks(object control) { return GetButtonImageCore("btnUpdateLinks"); }
        public IPictureDisp GetImage_btnLinkToExcel(object control) { return GetButtonImageCore("btnLinkToExcel"); }
        public IPictureDisp GetImage_btnViewLinks(object control) { return GetButtonImageCore("btnViewLinks"); }
        public IPictureDisp GetImage_btnSelectAllTables(object control) { return GetButtonImageCore("btnSelectAllTables"); }
        public IPictureDisp GetImage_btnFeedback(object control) { return GetButtonImageCore("btnFeedback"); }

        /// <summary>统一的图标分发：优先用配置里的自定义图片，没有则按配置回落内置图标。</summary>
        private static IPictureDisp GetButtonImageCore(string id)
        {
            try
            {
                IPictureDisp custom = RibbonConfig.GetImage(id);
                if (custom != null) return custom;
                string kind = RibbonConfig.BuiltinIconKind(id);
                if (kind != RibbonConfig.BuiltinUpdate && kind != RibbonConfig.BuiltinPaste
                    && kind != RibbonConfig.BuiltinView && kind != RibbonConfig.BuiltinMail && kind != RibbonConfig.BuiltinTable && kind != RibbonConfig.BuiltinFormula && kind != RibbonConfig.BuiltinFilter && kind != RibbonConfig.BuiltinLink && kind != RibbonConfig.BuiltinDirectory && kind != RibbonConfig.BuiltinAgent && kind != RibbonConfig.BuiltinWeb && kind != RibbonConfig.BuiltinDocument && kind != RibbonConfig.BuiltinSettings)
                    kind = RibbonConfig.BuiltinUpdate;
                return MakeRibbonIcon(kind, id == "btnFeedback" ? Ui.Amber : Ui.Blue, RibbonConfig.GetImageSize(id));
            }
            catch (Exception ex)
            {
                Log("GetButtonImageCore(" + (id ?? "") + ") 异常：" + ex.Message);
                return null;
            }
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
                        "WordUpdater", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch { }
            }
        }

        private static bool _idFailLogged;

        private static string ControlId(object control)
        {
            // 2026-09-26 实测根因：打开【网络盘/云同步映射盘】上的 docx 时，
            // WPS 传给回调的 control 代理对象读 Id 会失败或返回空——于是
            // getImage 拿到空 id 返回 null（图标白）、onAction 拿到空 id
            // 找不到网址（点了没反应）。本地文件完全正常。
            // 这里多试几种读法；都失败就把真实异常记进日志（只记首次防刷屏）。
            Exception last = null;
            try
            {
                dynamic c = control;
                string id = (string)c.Id;
                if (!string.IsNullOrEmpty(id)) return id;
            }
            catch (Exception ex) { last = ex; }
            try
            {
                object v = control.GetType().InvokeMember("Id",
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null, control, new object[0]);
                string id = v as string;
                if (!string.IsNullOrEmpty(id)) return id;
            }
            catch (Exception ex) { last = ex; }
            if (!_idFailLogged)
            {
                _idFailLogged = true;
                string tname;
                try { tname = control.GetType().FullName; }
                catch { tname = "(unknown)"; }
                Log("ControlId 读取失败（control 类型 " + tname + "）："
                    + (last == null ? "Id 返回空串" : last.GetType().Name + " - " + last.Message));
            }
            return "";
        }

        public void OnUpdateLinks(object control)
        {
            if (_busy) return;
            _busy = true;
            FlashSession flash = null;
            try
            {
                if (!EnsureApplication())
                {
                    UiDialog.Show(
                        "未检测到正在运行的 WPS 文字。\r\n请先打开 WPS 文字文档，再点击「更新链接」。",
                        "更新链接", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // ① 先清掉上一轮还没还原的标黄（防御：万一上次的 Timer 没跑完）
                ClearPendingFlash();

                RunResult result = UpdateEngine.Execute(_app);
                Log(string.Format(
                    "OnUpdateLinks -> {0} found={1} ok={2} fail={3} changed={4} notext={5} cleaned={6}",
                    result.Status, result.Found, result.Success, result.Failed,
                    result.Changes.Count, result.NoTextCount, result.CleanedHiddenHosts));

                // ② 引擎在各自更新完成的那一瞬间已把黄打在【可见 Range】上，
                //    这里只把 (Range, 原值) 收进 session —— 必须在弹窗【之前】做，
                //    此刻 COM 引用最新鲜，弹窗之后再取可能已被编辑或失效。
                flash = new FlashSession();
                for (int i = 0; i < result.Changes.Count; i++)
                {
                    LinkChangeInfo c = result.Changes[i];
                    if (c.FlashRange != null) flash.Add(c.FlashRange, c.PrevHighlight);
                }
                _flash = flash;

                // ③ 模态结果框。这段时间内 2.5 秒计时【没有】在走——
                //    用户点“确定”关掉弹窗的那一刻才开始（见下面的 finally）。
                try
                {
                    UiDialog.Show(FormatResult(result), "更新链接",
                        MessageBoxButtons.OK,
                        result.Status == RunStatus.Error ? MessageBoxIcon.Error : MessageBoxIcon.Information);
                }
                catch { }

                // ④ 弹窗已关闭：把视图滚到第一个变动处（不 Select、不 Activate，
                //    不抢选区也不抢焦点），让用户视线正好落在黄开始闪的地方。
                if (result.Changes.Count > 0 && result.Changes[0].FlashRange != null)
                    FlashHighlighter.ScrollIntoView(_app, result.Changes[0].FlashRange);
            }
            catch (Exception ex)
            {
                Log("OnUpdateLinks 异常：" + ex);
                try
                {
                    UiDialog.Show("发生未预期的错误：" + ex.Message, "更新链接",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
            }
            finally
            {
                // ⑤ 保命：无论弹窗正常关闭、还是中途抛异常/提前 return，
                //    都必须把计时开起来。漏了 = 黄色永久留在文档里，
                //    这是本功能最严重的失败模式。
                try { if (flash != null) flash.Start(2500); } catch { }
                _busy = false;
            }
        }

        // --------------------------- 「查看链接」回调 ---------------------------

        /// <summary>
        /// 打开「查看链接」窗口：列出当前文档引用的全部外部文件及其工作表/区域，
        /// 并可当场选择「更新所有链接」或「中断所有链接」。
        /// 与「更新链接」不同：这里扫描的是整篇文档，不需要先选中内容。
        ///
        /// 注意：这里必须用【非模态】Show() 而不是 ShowDialog() ——
        /// 模态弹窗期间 WPS 处于"对话框打开"状态，此时通过系统外壳打开
        /// xlsx（ShellExecute/DDE）会被 WPS 自己拒绝，报
        /// "对话框打开时命令无法执行"。改成非模态浮动窗就没有这个限制，
        /// 还能一边看着链接清单一边浏览文档。
        /// </summary>
        public void OnViewLinks(object control)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                if (!EnsureApplication())
                {
                    UiDialog.Show(
                        "未检测到正在运行的 WPS 文字。\r\n请先打开要检查的文档，再点击「查看链接」。",
                        "查看链接", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 已经开着就把它带到前台，不重复开窗
                if (_viewForm != null && !_viewForm.IsDisposed)
                {
                    try { _viewForm.Activate(); }
                    catch { }
                    return;
                }

                Log("OnViewLinks：打开查看链接窗口");
                try
                {
                    LinkViewResult probe = LinkViewEngine.Scan(_app);
                    Log(string.Format("OnViewLinks -> {0} links={1} sources={2}",
                        probe.Status, probe.TotalLinks, probe.Sources.Count));
                }
                catch (Exception scanEx)
                {
                    Log("OnViewLinks 预扫描异常：" + scanEx.Message);   // 只记录，不影响开窗
                }

                LinkViewForm form = new LinkViewForm(_app);
                form.FormClosed += delegate { _viewForm = null; };
                _viewForm = form;
                IntPtr owner = GetHostWindowHandle();
                if (owner != IntPtr.Zero)
                {
                    // Show(IWin32Window) 重载会在 Win32 层建立附属关系：
                    // 窗口一直浮在 WPS 之上、随它最小化，但不锁死它、也不压别的程序
                    // （切到 WPS 表格看数据时，列表会乖乖让位）。
                    form.Show(new WindowWrapper(owner));
                }
                else
                {
                    form.TopMost = true;   // 拿不到宿主窗口时的兜底，避免被主窗压住看不见
                    form.Show();
                }
                Log("OnViewLinks：窗口已打开（非模态）");
            }
            catch (Exception ex)
            {
                Log("OnViewLinks 异常：" + ex);
                try
                {
                    UiDialog.Show("发生未预期的错误：" + ex.Message, "查看链接",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
            }
            finally
            {
                _busy = false;
            }
        }

        // --------------------------- 「全选表格」回调 ---------------------------

        /// <summary>
        /// 选中当前文档中的全部表格（不含正文）。
        /// Word/WPS 的对象模型不支持"多段非连续选区"（Selection 只能是单段连续
        /// 区域），所以交给 TableSelectEngine 模拟人工的 Ctrl+框选：
        /// 逐个滚动到表格 → 取屏幕坐标 → 按住 Ctrl 模拟鼠标拖框。
        /// 选完后用户可以接着用功能区的表格工具 / 字体命令统一处理。
        /// </summary>
        public void OnSelectAllTables(object control)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                if (!EnsureApplication())
                {
                    UiDialog.Show(
                        "未检测到正在运行的 WPS 文字。\r\n请先打开文档，再点击「全选表格」。",
                        "全选表格", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                int total, selected;
                string err = TableSelectEngine.SelectAllTables(_app, out total, out selected);
                Log(string.Format("OnSelectAllTables -> total={0} selected={1} err={2}",
                    total, selected, err ?? "(null)"));
                if (err != null)
                {
                    UiDialog.Show(err, "全选表格",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (total == 0)
                {
                    UiDialog.Show("本文档的正文中没有表格。", "全选表格",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (selected < total)
                {
                    UiDialog.Show(string.Format(
                        "全文共 {0} 张表格，成功选中 {1} 张。\r\n"
                        + "有 {2} 张未能定位（可能被隐藏或在受保护区域），可手工 Ctrl 补选。",
                        total, selected, total - selected),
                        "全选表格", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                // 全部选中时不弹窗：选区本身就是反馈，弹窗反而打断后续操作
            }
            catch (Exception ex)
            {
                Log("OnSelectAllTables 异常：" + ex);
                try
                {
                    UiDialog.Show("发生未预期的错误：" + ex.Message, "全选表格",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
            }
            finally
            {
                _busy = false;
            }
        }

        /// <summary>卸载/关文档时关掉还开着的查看窗口，避免它挂在已死的宿主上。</summary>
        public static void CloseViewForm()
        {
            LinkViewForm f = _viewForm;
            _viewForm = null;
            if (f == null) return;
            try
            {
                if (!f.IsDisposed && f.IsHandleCreated)
                    f.BeginInvoke((MethodInvoker)delegate { try { f.Close(); } catch { } });
            }
            catch
            {
                try { f.Close(); }
                catch { }
            }
        }

        /// <summary>尽力拿 WPS 主窗口句柄，给非模态窗当 Owner 用。</summary>
        private static IntPtr GetHostWindowHandle()
        {
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

        // --------------------------- 「链接Excel」回调 ---------------------------

        public void OnLinkToExcel(object control)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                if (!EnsureApplication())
                {
                    UiDialog.Show(
                        "未检测到正在运行的 WPS 文字。\r\n请先打开 WPS 文字文档，再点击「链接Excel」。",
                        "链接Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                LinkPasteResult result = LinkPasteEngine.Execute(_app);
                Log(string.Format("OnLinkToExcel -> {0} linked={1} refreshed={2} rolledBack={3} src={4}",
                    result.Status, result.CellsLinked, result.Refreshed, result.RolledBack,
                    result.SourceDesc));
                MessageBoxIcon icon = result.Status == LinkPasteStatus.Error
                    ? MessageBoxIcon.Error
                    : (result.Status == LinkPasteStatus.Completed
                        ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
                UiDialog.Show(FormatLinkResult(result), "链接Excel",
                    MessageBoxButtons.OK, icon);
            }
            catch (Exception ex)
            {
                Log("OnLinkToExcel 异常：" + ex);
                try
                {
                    UiDialog.Show("发生未预期的错误：" + ex.Message, "链接Excel",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
            }
            finally
            {
                _busy = false;
            }
        }

        // --------------------------- 链接粘贴结果文案 ---------------------------

        private static string FormatLinkResult(LinkPasteResult r)
        {
            switch (r.Status)
            {
                case LinkPasteStatus.NoDocument:
                    return "当前没有打开的文档。";
                case LinkPasteStatus.NotInTable:
                    return "光标未位于表格内。\r\n\r\n请先点击到要链接的文档表格中，再点击「链接Excel」。";
                case LinkPasteStatus.TableNotUniform:
                    return "所选区域包含合并单元格、嵌套表格或不规则结构，无法与源区域逐格对应。\r\n\r\n请在同一张表格内选择由独立单元格组成的规则矩形区域后重试。";
                case LinkPasteStatus.TableTooLarge:
                    return string.Format("选中的区域共 {0} 行 × {1} 列 = {2} 个单元格，超出单次支持上限（{3} 个）。\r\n请缩小选中的范围后重试。",
                        r.TargetRows, r.TargetCols, r.TargetCells, LinkPasteEngine.MaxCells);
                case LinkPasteStatus.ClipboardInvalid:
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("剪贴板内容不是可链接的单元格区域。\r\n\r\n请先在 WPS 表格中复制要链接的单元格区域（可按住 Ctrl 多选多个区域），再点击「链接Excel」。");
                        if (!string.IsNullOrEmpty(r.ErrorMessage))
                        {
                            sb.Append("\r\n\r\n技术详情：");
                            sb.Append(r.ErrorMessage);
                        }
                        return sb.ToString();
                    }
                case LinkPasteStatus.MultiArea:
                    return "复制的是多个不连续的单元格区域，但剪贴板没能还原它们的明细位置。\r\n\r\n请保持 WPS 表格处于打开状态、且当前选区仍是刚复制的那几个格子，然后点击「链接Excel」重试。";
                case LinkPasteStatus.SizeMismatch:
                    {
                        bool overflow = r.SourceCells > r.TargetCells;
                        string head = overflow
                            ? "源数据的格数多于选中的区域（区域溢出）。"
                            : "源数据的格数少于选中的区域（区域未填满）。";
                        string tail = overflow
                            ? "请扩大选中的区域，或改复制更小的区域后重试。"
                            : "请缩小选中的区域，或补齐复制的格子后重试。";
                        return string.Format(
                            head + "\r\n\r\n复制的源：{0} 个单元格\r\n选中的区域：{1} 行 × {2} 列 = {3} 个单元格"
                            + "\r\n\r\n链接无效，文档中未作任何修改。\r\n" + tail,
                            r.SourceCells, r.TargetRows, r.TargetCols, r.TargetCells);
                    }
                case LinkPasteStatus.SelectionEmpty:
                    return "没有选中任何表格单元格。\r\n\r\n请先在表格里拖选出要链接的矩形区域"
                        + "（要链接整张表格，就把整张表都选上），再点击「链接Excel」。";
                case LinkPasteStatus.SelectionNotRect:
                    return "所选区域包含合并单元格、嵌套表格或不规则结构，无法逐格对应。\r\n\r\n" + (r.ErrorMessage ?? "")
                        + "\r\n\r\n请重新选择由独立单元格组成、列边界对齐的连续矩形区域后重试。";
                case LinkPasteStatus.SelectionOutOfTable:
                    return "选区超出了表格范围。\r\n\r\n" + (r.ErrorMessage ?? "")
                        + "\r\n\r\n请只在同一张表格内部拖选矩形区域（不要拖到表格外的正文，也不要跨表）。";
                case LinkPasteStatus.SourceMissing:
                    return "源工作簿文件不存在或已被移动：\r\n\r\n" + (r.SourceDesc ?? "") +
                        "\r\n\r\n请确认源文件位置后重试。";
                case LinkPasteStatus.Error:
                    return "链接粘贴未能完成" +
                        (r.RolledBack ? "（已自动撤销已写入的部分，未产生残缺链接）。\r\n" : "。\r\n") +
                        "错误信息：" + (r.ErrorMessage ?? "未知");
                case LinkPasteStatus.Completed:
                    {
                        string tpl = RibbonConfig.GetResultText("OnLinkToExcel");
                        if (tpl != null)
                        {
                            // 用户在 配置器脚本 里自定义了完成文案
                            StringBuilder t = new StringBuilder();
                            t.Append(FillPasteTemplate(tpl, r));
                            if (r.Failures.Count > 0)
                            {
                                // 刷新失败明细是安全信息，自定义模板也照附（配置里说明过）
                                t.AppendLine();
                                t.AppendLine();
                                t.AppendLine("刷新失败明细：");
                                int maxF = Math.Min(r.Failures.Count, 5);
                                for (int i = 0; i < maxF; i++)
                                    t.Append("  " + r.Failures[i] + "\r\n");
                                if (r.Failures.Count > maxF)
                                    t.AppendFormat("  …… 其余 {0} 条略。", r.Failures.Count - maxF);
                            }
                            if (r.CleanedHiddenHosts > 0 &&
                                tpl.IndexOf("{cleaned}", StringComparison.Ordinal) < 0)
                            {
                                t.AppendLine();
                                t.AppendFormat("（已自动释放 {0} 个刷新时拉起的后台 WPS 表格进程，源文件占用已解除。）",
                                    r.CleanedHiddenHosts);
                            }
                            return t.ToString();
                        }

                        StringBuilder sb = new StringBuilder();
                        sb.AppendFormat("链接粘贴完成：已在选中的 {0} 行 × {1} 列里为 {2} 个单元格建立链接",
                            r.TargetRows, r.TargetCols, r.CellsLinked);
                        if (r.Refreshed >= r.CellsLinked && r.Failures.Count == 0)
                        {
                            sb.Append("，并已全部刷新。");
                        }
                        else
                        {
                            sb.AppendFormat("，其中 {0} 个已刷新。", r.Refreshed);
                            if (r.Failures.Count > 0)
                            {
                                sb.Append("\r\n其余刷新失败（链接本身有效，源数据可用时可选中表格点「更新链接」重试）。");
                            }
                        }
                        sb.AppendLine();
                        sb.AppendLine();
                        sb.Append("源：");
                        sb.Append(string.IsNullOrEmpty(r.SourceDesc) ? "(未知)" : r.SourceDesc);
                        sb.AppendLine();
                        sb.Append("目标：");
                        if (r.TargetRows == 1 && r.TargetCols == 1)
                            sb.Append("第 " + r.TargetRow1 + " 行第 " + r.TargetCol1 + " 列");
                        else
                            sb.Append("第 " + r.TargetRow1 + "~" + (r.TargetRow1 + r.TargetRows - 1)
                                + " 行 × 第 " + r.TargetCol1 + "~" + (r.TargetCol1 + r.TargetCols - 1) + " 列");
                        sb.AppendLine();
                        sb.AppendLine();
                        sb.Append("提示：源数据变动后，选中该表格并点击「更新链接」即可一键同步。");
                        if (r.Failures.Count > 0)
                        {
                            sb.AppendLine();
                            sb.AppendLine();
                            sb.AppendLine("刷新失败明细：");
                            int max = Math.Min(r.Failures.Count, 5);
                            for (int i = 0; i < max; i++)
                                sb.Append("  " + r.Failures[i] + "\r\n");
                            if (r.Failures.Count > max)
                                sb.AppendFormat("  …… 其余 {0} 条略。", r.Failures.Count - max);
                        }
                        if (r.CleanedHiddenHosts > 0)
                        {
                            sb.AppendLine();
                            sb.AppendFormat("（已自动释放 {0} 个刷新时拉起的后台 WPS 表格进程，源文件占用已解除。）",
                                r.CleanedHiddenHosts);
                        }
                        return sb.ToString();
                    }
                default:
                    return "未知状态。";
            }
        }

        // --------------------------- 结果文案 ---------------------------

        /// <summary>把更新链接的自定义模板里的占位符换成实际数字。</summary>
        private static string FillUpdateTemplate(string tpl, RunResult r)
        {
            return tpl
                .Replace("{found}", r.Found.ToString())
                .Replace("{success}", r.Success.ToString())
                .Replace("{failed}", r.Failed.ToString())
                .Replace("{changed}", r.Changes.Count.ToString())
                .Replace("{cleaned}", r.CleanedHiddenHosts.ToString());
        }

        /// <summary>把链接Excel的自定义模板里的占位符换成实际数字/文本。</summary>
        private static string FillPasteTemplate(string tpl, LinkPasteResult r)
        {
            return tpl
                .Replace("{rows}", r.TargetRows.ToString())
                .Replace("{cols}", r.TargetCols.ToString())
                .Replace("{cells}", r.CellsLinked.ToString())
                .Replace("{refreshed}", r.Refreshed.ToString())
                .Replace("{failures}", r.Failures.Count.ToString())
                .Replace("{source}", string.IsNullOrEmpty(r.SourceDesc) ? "(未知)" : r.SourceDesc)
                .Replace("{cleaned}", r.CleanedHiddenHosts.ToString());
        }

        private static string FormatResult(RunResult r)
        {
            switch (r.Status)
            {
                case RunStatus.NoDocument:
                    return "当前没有打开的文档。";
                case RunStatus.NoSelection:
                    return "请先选中包含链接的内容范围（按住鼠标拖选，或按 Ctrl+A 全选），再点击「更新链接」。";
                case RunStatus.NoLinks:
                    return "所选内容中未发现任何链接。\r\n\r\n提示：从 Excel 复制单元格后，在 WPS 文字中「粘贴」→「选择性粘贴」→「粘贴链接」，即可创建可更新的链接。";
                case RunStatus.Error:
                    return "更新过程中发生错误：\r\n" + (r.ErrorMessage ?? "未知错误");
                case RunStatus.Completed:
                    {
                        StringBuilder sb = new StringBuilder();
                        string tpl = RibbonConfig.GetResultText("OnUpdateLinks");
                        if (tpl != null)
                        {
                            // 用户在 配置器脚本 里自定义了完成文案
                            sb.Append(FillUpdateTemplate(tpl, r));
                        }
                        else
                        {
                            sb.AppendFormat("更新完成：共发现 {0} 个链接，成功 {1} 个，数值变动 {2} 个",
                                r.Found, r.Success, r.Changes.Count);
                            if (r.Failed > 0)
                            {
                                sb.AppendFormat("，失败 {0} 个。", r.Failed);
                            }
                            else
                            {
                                sb.Append("。");
                            }
                        }

                        // 变动明细：哪些链接的数值变了（这些位置已标黄）
                        if (r.Changes.Count > 0)
                        {
                            sb.AppendLine();
                            sb.AppendLine();
                            sb.AppendLine("变动明细（已标黄，关闭本窗口后 2.5 秒自动还原）：");
                            int max = Math.Min(r.Changes.Count, 10);
                            for (int i = 0; i < max; i++)
                            {
                                LinkChangeInfo c = r.Changes[i];
                                string oldText = c.OldText.Length == 0 ? "（空）" : c.OldText;
                                string newText = c.NewText.Length == 0 ? "（空）" : c.NewText;
                                sb.AppendFormat("  {0}. {1}：{2} → {3}\r\n",
                                    c.Index, c.Where, oldText, newText);
                            }
                            if (r.Changes.Count > max)
                                sb.AppendFormat("  …… 其余 {0} 条变动略。", r.Changes.Count - max);
                        }

                        if (r.Failed > 0)
                        {
                            // 失败明细是安全信息，自定义模板也照附（配置里说明过）
                            sb.AppendLine();
                            sb.AppendLine();
                            sb.AppendLine("失败明细：");
                            int max = Math.Min(r.Failures.Count, 10);
                            for (int i = 0; i < max; i++)
                            {
                                LinkFailureInfo f = r.Failures[i];
                                string src = string.IsNullOrEmpty(f.Source) ? "(未知)" : f.Source;
                                sb.AppendFormat("  {0}. [{1}] 源：{2}\r\n     原因：{3}\r\n",
                                    f.Index, f.Kind, src, f.Reason);
                            }
                            if (r.Failures.Count > max)
                                sb.AppendFormat("  …… 其余 {0} 条失败信息略。", r.Failures.Count - max);
                        }
                        // 更新成功但没有可比对文本的：链接图片/对象（或读不到文本）
                        if (r.NoTextCount > 0)
                        {
                            sb.AppendLine();
                            sb.AppendFormat("其中 {0} 个为图片/对象，无文本可比。", r.NoTextCount);
                        }
                        if (r.CleanedHiddenHosts > 0 &&
                            (tpl == null || tpl.IndexOf("{cleaned}", StringComparison.Ordinal) < 0))
                        {
                            sb.AppendLine();
                            sb.AppendFormat("（已自动释放 {0} 个更新时拉起的后台 WPS 表格进程，源文件占用已解除。）",
                                r.CleanedHiddenHosts);
                        }
                        return sb.ToString();
                    }
                default:
                    return "未知状态。";
            }
        }

        // --------------------------- 按钮图标 ---------------------------

        private static IPictureDisp MakeRibbonIcon(string kind, Color accent, int logicalSize)
        {
            return RibbonImages.Create(logicalSize,
                delegate(int width, int height) { return UiIcons.CreateRibbon(kind, accent, width, height); }, Log);
        }

        // --------------------------- 日志 ---------------------------

        private static readonly object _logLock = new object();

        public static void Log(string msg)
        {
            try
            {
                // 固定日志位置：VSTO 加载时 Assembly.Location 指向 CLR 的
                // dl3 影子缓存而非部署目录，因此写到 %LocalAppData% 下。
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordUpdater");
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "WordUpdater.log");
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

        private static string ComTypeName(object o)
        {
            if (o == null) return "(null)";
            try { return o.GetType().Name; }
            catch { return "?"; }
        }
    }
}
