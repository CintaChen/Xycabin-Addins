using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

// 功能：Excel-Iterate → 数据 → 审计工具箱 → 表头匹配
// 控件：btnMergeHeaders
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnMergeHeaders(object control)
        {
            if (_busy) return;
            if (_mergeForm != null && !_mergeForm.IsDisposed) { _mergeForm.Activate(); return; }
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。");
                MergeHeadersForm form = new MergeHeadersForm();
                ApplyActiveTarget(form);
                form.PreviewRequested += delegate { RunPreview(form); };
                form.ExecuteRequested += delegate { RunExecute(form); };
                form.UndoRequested += delegate { RunUndo(form); };
                form.ConfirmRequested += delegate { RunConfirm(form); };
                form.TargetSheetsRequested += delegate(string path) { RunLoadSheets(form, path, false); };
                form.SourceSheetsRequested += delegate(string path) { RunLoadSheets(form, path, true); };
                form.FormClosed += delegate
                {
                    _mergeForm = null;
                    try { if (_mergeGateway != null) { _mergeGateway.Dispose(); } }
                    catch (Exception ex) { Log("表头匹配合并：释放工作簿引用失败 " + ex.Message); }
                    finally { _mergeGateway = null; _mergeTarget = null; _mergeSnapshot = null; _mergeBackupPath = null; }
                    form.Dispose();
                };
                _mergeForm = form;
                // 与链接助手/按列筛选等非模态弹窗完全一致：取宿主主窗口作 owner 显示。
                IntPtr owner = GetHostWindowHandle();
                Log("表头匹配合并：打开前 enabled=" + HostDialogService.HostEnabled()
                    + " / 前台=" + NativeOle.GetForegroundWindow() + " / owner=" + owner);
                if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
                Log("表头匹配合并：Show 后 enabled=" + HostDialogService.HostEnabled()
                    + " / 前台=" + NativeOle.GetForegroundWindow()
                    + " / 弹窗=" + (form.IsHandleCreated ? form.Handle : IntPtr.Zero));
            }
            catch (Exception ex)
            {
                Log("表头匹配合并：" + ex);
                UiDialog.Show(ex.Message, "跨工作簿表头匹配合并", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        public IPictureDisp GetImage_btnMergeHeaders(object control)
        { return GetButtonImageCore("btnMergeHeaders"); }
    }
}
