using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 生成目录
// 控件：btnGenerateDirectory
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnGenerateDirectory(object control)
        {
            if (_busy) return; _busy = true;
            try { if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。"); if (_directoryNavigation == null) _directoryNavigation = new DirectoryNavigation(_app); int count = DirectoryNavigation.Generate(_app); ShowStatus("生成目录：已链接 " + count + " 张工作表"); }
            catch (Exception ex) { Log("生成目录：" + ex); UiDialog.Show(ex.Message, "生成目录", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { _busy = false; }
        }

        public IPictureDisp GetImage_btnGenerateDirectory(object control) { return GetButtonImageCore("btnGenerateDirectory"); }
    }
}
