using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 一键公式 → 数值计算 → 乘 → 自定义计算
// 控件：btnCustomCalculateMultiply
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnCustomCalculateMultiply(object control) { RunFormula(FormulaOperation.Multiply, true); }

        public IPictureDisp GetImage_btnCustomCalculateMultiply(object control) { return GetButtonImageCore("btnCustomCalculateMultiply"); }
    }
}
