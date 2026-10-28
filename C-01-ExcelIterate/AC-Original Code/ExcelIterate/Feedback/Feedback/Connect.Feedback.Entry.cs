using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 反馈 → 问题反馈
// 控件：btnFeedback
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnOpenUrl_btnFeedback(object control) { OpenUrlCore("btnFeedback"); }

        public IPictureDisp GetImage_btnFeedback(object control)
        { return GetButtonImageCore("btnFeedback"); }
    }
}
