using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 一键公式 → 移除公式封装
// 控件：btnUnwrapFormula
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnUnwrapFormula(object control) { RunFormula(FormulaOperation.Unwrap); }

        public IPictureDisp GetImage_btnUnwrapFormula(object control) { return GetButtonImageCore("btnUnwrapFormula"); }
    }
}
