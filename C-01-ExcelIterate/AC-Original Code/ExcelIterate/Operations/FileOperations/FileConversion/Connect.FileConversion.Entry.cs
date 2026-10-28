using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Excel-Iterate → 操作 → 文件操作 → 文件转换与 PDF 工具
// 控件：menuFileConversion
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuFileConversion(object control) { return GetButtonImageCore("menuFileConversion"); }
    }
}
