using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

// 功能：WordUpdater → 链接 → 查看链接
// 控件：btnViewLinks
namespace WordUpdater
{
    public partial class Connect
    {
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

        public IPictureDisp GetImage_btnViewLinks(object control) { return GetButtonImageCore("btnViewLinks"); }
    }
}
