using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 一键公式 → 数值计算 → 除 → 默认除
// 控件：btnCalculateDivide
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnCalculateDivide(object control) { RunFormula(FormulaOperation.Divide); }

        public IPictureDisp GetImage_btnCalculateDivide(object control) { return GetButtonImageCore("btnCalculateDivide"); }
    }
}
