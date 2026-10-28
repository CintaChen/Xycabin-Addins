using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Excel-Iterate → 操作 → 文件操作 → 文件转换与 PDF 工具 → 批量文件转 PDF
// 控件：btnConvertPdf
// 共用实现：@shared/src/addin/Connect.Pdf.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnConvertPdf(object control) { OpenPdfTool("convert"); }

        public IPictureDisp GetImage_btnConvertPdf(object control) { return GetButtonImageCore("btnConvertPdf"); }
    }
}
