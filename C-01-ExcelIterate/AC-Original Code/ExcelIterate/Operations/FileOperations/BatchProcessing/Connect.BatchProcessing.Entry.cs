using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Excel-Iterate → 操作 → 文件操作 → 批量处理
// 控件：menuBatchProcessing
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuBatchProcessing(object control) { return GetButtonImageCore("menuBatchProcessing"); }
    }
}
