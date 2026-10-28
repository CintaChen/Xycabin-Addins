using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：WordUpdater → 操作 → 文件操作 → 文件转换与 PDF 工具 → PDF 拆分
// 控件：btnSplitPdf
// 共用实现：@shared/src/addin/Connect.Pdf.cs
namespace WordUpdater
{
    public partial class Connect
    {
        public void OnSplitPdf(object control) { OpenPdfTool("split"); }

        public IPictureDisp GetImage_btnSplitPdf(object control) { return GetButtonImageCore("btnSplitPdf"); }
    }
}
