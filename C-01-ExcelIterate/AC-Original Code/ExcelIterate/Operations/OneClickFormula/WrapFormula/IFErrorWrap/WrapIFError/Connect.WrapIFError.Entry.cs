using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 一键公式 → 进行公式封装 → IFERROR封装 → 默认返回空
// 控件：btnWrapIFError
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnWrapIFError(object control) { RunFormula(FormulaOperation.IFError); }

        public IPictureDisp GetImage_btnWrapIFError(object control) { return GetButtonImageCore("btnWrapIFError"); }
    }
}
