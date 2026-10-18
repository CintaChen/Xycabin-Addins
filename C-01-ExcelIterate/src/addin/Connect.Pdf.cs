using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Excel_Iterate
{
    public partial class Connect
    {
        private static readonly Dictionary<string, PdfToolsForm> pdfForms = new Dictionary<string, PdfToolsForm>();
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
                var form = new PdfToolsForm(mode); pdfForms[mode] = form;
                form.FormClosed += delegate { if (pdfForms.TryGetValue(mode, out current) && ReferenceEquals(current, form)) pdfForms.Remove(mode); form.Dispose(); };
                ShowFileForm(form);
            }
            catch (Exception error) { Log("PDF 工具：" + error); UiDialog.Show(error.Message, PdfToolsForm.Title(mode), MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        internal static void ClosePdfTools()
        { foreach (var form in new List<PdfToolsForm>(pdfForms.Values)) if (!form.IsDisposed) form.Dispose(); pdfForms.Clear(); }
    }
}
