using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 数据 → 筛选增强 → 快捷高筛
// 控件：btnQuickAdvancedFilter
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnQuickAdvancedFilter(object control)
        {
            if (_busy) return; _busy = true;
            try { if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。"); int count = QuickAdvancedFilter.Run(_app); ShowStatus("快捷高筛：已输出 " + count + " 条记录至新工作表"); }
            catch (Exception ex) { Log("快捷高筛：" + ex); UiDialog.Show(ex.Message, "快捷高筛", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { _busy = false; }
        }

        public IPictureDisp GetImage_btnQuickAdvancedFilter(object control) { return GetButtonImageCore("btnQuickAdvancedFilter"); }
    }
}
