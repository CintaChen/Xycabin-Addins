using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 文件操作
// 控件：menuFileOperations
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuFileOperations(object control) { return GetButtonImageCore("menuFileOperations"); }
    }
}
