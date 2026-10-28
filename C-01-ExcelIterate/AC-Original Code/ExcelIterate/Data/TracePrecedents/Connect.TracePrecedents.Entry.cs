using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 数据 → 追踪引用
// 控件：btnTracePrecedents
namespace Excel_Iterate
{
    public partial class Connect
    {
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

        public IPictureDisp GetImage_btnTracePrecedents(object control)
        { return GetTraceImage("btnTracePrecedents", false); }
    }
}
