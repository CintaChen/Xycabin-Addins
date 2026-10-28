using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 链接助手
// 控件：btnLinkAssistant
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnLinkAssistant(object control)
        {
            if (_busy) return;
            if (_linkForm != null && !_linkForm.IsDisposed) { _linkForm.Activate(); return; }
            try { if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。"); LinkAssistantForm form = new LinkAssistantForm(new LinkAssistantSession(_app)); _linkForm = form; form.FormClosed += delegate { _linkForm = null; form.Dispose(); }; IntPtr owner = GetHostWindowHandle(); if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner)); }
            catch (Exception ex) { Log("链接助手：" + ex); UiDialog.Show(ex.Message, "链接助手", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        public IPictureDisp GetImage_btnLinkAssistant(object control) { return GetButtonImageCore("btnLinkAssistant"); }
    }
}
