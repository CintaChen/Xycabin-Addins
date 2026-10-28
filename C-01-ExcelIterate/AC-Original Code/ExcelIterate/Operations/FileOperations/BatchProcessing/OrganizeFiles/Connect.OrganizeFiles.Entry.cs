using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 文件操作 → 批量处理 → 文件批量整理
// 控件：btnOrganizeFiles
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnOrganizeFiles(object control)
        {
            if (_busy) return;
            if (_organizeForm != null && !_organizeForm.IsDisposed) { _organizeForm.Activate(); return; }
            try { var form = new OrganizeFilesForm(); _organizeForm = form; form.FormClosed += delegate { _organizeForm = null; form.Dispose(); }; ShowFileForm(form); }
            catch (Exception ex) { Log("文件批量整理："+ex); UiDialog.Show(ex.Message,"文件批量整理",MessageBoxButtons.OK,MessageBoxIcon.Warning); }
        }

        public IPictureDisp GetImage_btnOrganizeFiles(object control) { return GetButtonImageCore("btnOrganizeFiles"); }
    }
}
