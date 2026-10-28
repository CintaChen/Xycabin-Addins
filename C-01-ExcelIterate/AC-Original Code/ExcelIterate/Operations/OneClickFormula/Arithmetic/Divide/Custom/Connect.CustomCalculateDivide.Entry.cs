using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 一键公式 → 数值计算 → 除 → 自定义计算
// 控件：btnCustomCalculateDivide
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnCustomCalculateDivide(object control) { RunFormula(FormulaOperation.Divide, true); }

        public IPictureDisp GetImage_btnCustomCalculateDivide(object control) { return GetButtonImageCore("btnCustomCalculateDivide"); }
    }
}
