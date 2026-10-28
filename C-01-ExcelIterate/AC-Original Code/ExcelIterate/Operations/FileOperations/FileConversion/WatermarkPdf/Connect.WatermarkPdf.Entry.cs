using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Excel-Iterate → 操作 → 文件操作 → 文件转换与 PDF 工具 → PDF 添加水印
// 控件：btnWatermarkPdf
// 共用实现：@shared/src/addin/Connect.Pdf.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnWatermarkPdf(object control) { OpenPdfTool("watermark"); }

        public IPictureDisp GetImage_btnWatermarkPdf(object control) { return GetButtonImageCore("btnWatermarkPdf"); }
    }
}
