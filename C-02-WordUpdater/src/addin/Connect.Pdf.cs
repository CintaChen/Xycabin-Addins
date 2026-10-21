using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    public partial class Connect
    {
        private static readonly Dictionary<string, PdfToolsForm> pdfForms = new Dictionary<string, PdfToolsForm>();
        public void OnCurrentPdf(object control) { OpenPdfTool("current"); }
        public IPictureDisp GetImage_btnCurrentPdf(object control) { return GetButtonImageCore("btnCurrentPdf"); }
        public IPictureDisp GetImage_menuBatchProcessing(object control) { return GetButtonImageCore("menuBatchProcessing"); }
        public IPictureDisp GetImage_menuFileConversion(object control) { return GetButtonImageCore("menuFileConversion"); }
#if WPS_WORD
        public IPictureDisp GetImage_menuFileOperations(object control) { return GetButtonImageCore("menuFileOperations"); }
#endif
        public void OnConvertPdf(object control) { OpenPdfTool("convert"); }
        public void OnWatermarkPdf(object control) { OpenPdfTool("watermark"); }
        public void OnSplitPdf(object control) { OpenPdfTool("split"); }
        public void OnMergePdf(object control) { OpenPdfTool("merge"); }
        public IPictureDisp GetImage_btnConvertPdf(object control) { return GetButtonImageCore("btnConvertPdf"); }
        public IPictureDisp GetImage_btnWatermarkPdf(object control) { return GetButtonImageCore("btnWatermarkPdf"); }
        public IPictureDisp GetImage_btnSplitPdf(object control) { return GetButtonImageCore("btnSplitPdf"); }
        public IPictureDisp GetImage_btnMergePdf(object control) { return GetButtonImageCore("btnMergePdf"); }
        private static void OpenPdfTool(string mode)
        {
            PdfToolsForm current; if (pdfForms.TryGetValue(mode, out current) && !current.IsDisposed) { current.Activate(); return; }
            try
            {
                PdfCurrentDocument session = null;
                if (mode == "current") { if (!EnsureApplication()) throw new InvalidOperationException("请先打开文档。"); ValidatePdfHost(); session = new PdfCurrentDocument(_app); }
                PdfToolsForm form;
                try { form = new PdfToolsForm(mode); }
                catch { if (session != null) session.Dispose(); throw; }
                if (session != null) { form.ExportCurrent = session.Export; form.CloseCurrent = session.Dispose; form.SetCurrentSource(session.DisplayName); }
                pdfForms[mode] = form;
                form.FormClosed += delegate { if (pdfForms.TryGetValue(mode, out current) && ReferenceEquals(current, form)) pdfForms.Remove(mode); form.Dispose(); };
                try { IntPtr owner = GetHostWindowHandle(); if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner)); }
                catch { pdfForms.Remove(mode); form.Dispose(); throw; }
            }
            catch (Exception error) { Log("PDF 工具：" + error); UiDialog.Show(error.Message, PdfToolsForm.Title(mode), MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        private static void ValidatePdfHost()
        {
            object window = null;
            try
            {
                window = ((dynamic)_app).ActiveWindow;
                uint process; GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)window).Hwnd)), out process);
                if (process != (uint)System.Diagnostics.Process.GetCurrentProcess().Id)
                    throw new InvalidOperationException("当前文档不属于本加载项所在的宿主实例，请在目标文档所在实例中重试。");
            }
            finally { if (window != null && Marshal.IsComObject(window)) Marshal.ReleaseComObject(window); }
        }
        internal static void ClosePdfTools()
        { foreach (var form in new List<PdfToolsForm>(pdfForms.Values)) if (!form.IsDisposed) form.Dispose(); pdfForms.Clear(); }
    }
}
