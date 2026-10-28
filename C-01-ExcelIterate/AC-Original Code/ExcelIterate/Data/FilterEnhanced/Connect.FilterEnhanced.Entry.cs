using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 数据 → 筛选增强
// 控件：menuFilterEnhanced
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuFilterEnhanced(object control) { return GetButtonImageCore("menuFilterEnhanced"); }
    }
}
