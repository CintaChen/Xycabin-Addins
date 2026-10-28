using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 一键公式 → 数值计算 → 加 → 默认加
// 控件：btnCalculateAdd
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnCalculateAdd(object control) { RunFormula(FormulaOperation.Add); }

        public IPictureDisp GetImage_btnCalculateAdd(object control) { return GetButtonImageCore("btnCalculateAdd"); }
    }
}
