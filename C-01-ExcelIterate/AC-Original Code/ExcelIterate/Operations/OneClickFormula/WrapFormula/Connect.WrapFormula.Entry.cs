using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 一键公式 → 进行公式封装
// 控件：menuWrapFormula
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuWrapFormula(object control) { return GetButtonImageCore("menuWrapFormula"); }
    }
}
