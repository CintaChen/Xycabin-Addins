using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

// 功能：WordUpdater → 链接 → 更新链接
// 控件：btnUpdateLinks
namespace WordUpdater
{
    public partial class Connect
    {
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

        public IPictureDisp GetImage_btnUpdateLinks(object control) { return GetButtonImageCore("btnUpdateLinks"); }
    }
}
