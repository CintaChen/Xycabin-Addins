using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 文件操作 → 批量处理 → 文件批量命名
// 控件：btnRenameFiles
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnRenameFiles(object control)
        {
            if (_busy) return;
            if (_renameForm != null && !_renameForm.IsDisposed) { _renameForm.Activate(); return; }
            try { if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。"); var form = new RenameFilesForm(new RenameWorkbookSession(_app)); _renameForm = form; form.FormClosed += delegate { _renameForm = null; form.Dispose(); }; ShowFileForm(form); }
            catch (Exception ex) { Log("文件批量命名："+ex); UiDialog.Show(ex.Message,"文件批量命名",MessageBoxButtons.OK,MessageBoxIcon.Warning); }
        }

        public IPictureDisp GetImage_btnRenameFiles(object control) { return GetButtonImageCore("btnRenameFiles"); }
    }
}
