using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 数据 → 筛选增强 → 按列筛选
// 控件：btnFilterColumns
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnFilterColumns(object control)
        {
            if (_busy) return;
            if (_columnForm != null && !_columnForm.IsDisposed) { _columnForm.Activate(); return; }
            _busy = true;
            ComColumnVisibilityHost host = null;
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格并选中一行。");
                host = new ComColumnVisibilityHost(_app);
                var columns = host.Collect();
                ColumnFilterForm form = new ColumnFilterForm(columns, host.RangeDescription);
                ComColumnVisibilityHost bound = host;
                form.ApplyRequested += delegate(System.Collections.Generic.IList<ColumnChoice> choices)
                {
                    if (_busy) return;
                    _busy = true;
                    try
                    {
                        FormulaCommandResult result = ColumnVisibilityEngine.Apply(bound, choices,
                            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Excel_Iterate", "ColumnRecovery"));
                        if (!result.Failed)
                        {
                            if (result.Changed > 0) ShowStatus("按列筛选：已修改 " + result.Changed + " 列的显示状态");
                            form.Close();
                        }
                        else { Log(result.Details); UiDialog.Show(form, result.Details, "按列筛选", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                    }
                    finally { _busy = false; }
                };
                form.FormClosed += delegate { _columnForm = null; try { bound.Dispose(); } catch (Exception ex) { Log("列筛选引用释放：" + ex.Message); } form.Dispose(); };
                _columnForm = form;
                IntPtr owner = GetHostWindowHandle();
                if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
                host = null; // The non-modal form owns the snapshot until it closes.
            }
            catch (Exception ex) { Log("按列筛选：" + ex); UiDialog.Show(ex.Message, "按列筛选", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { if (host != null) host.Dispose(); _busy = false; }
        }

        public IPictureDisp GetImage_btnFilterColumns(object control) { return GetButtonImageCore("btnFilterColumns"); }
    }
}
