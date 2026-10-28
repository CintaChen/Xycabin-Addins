using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：WordUpdater → 操作 → 文件操作 → 文件转换与 PDF 工具 → PDF 合并
// 控件：btnMergePdf
// 共用实现：@shared/src/addin/Connect.Pdf.cs
namespace WordUpdater
{
    public partial class Connect
    {
        public void OnMergePdf(object control) { OpenPdfTool("merge"); }

        public IPictureDisp GetImage_btnMergePdf(object control) { return GetButtonImageCore("btnMergePdf"); }
    }
}
