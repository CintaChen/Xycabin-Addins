using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

// 功能：WordUpdater → 操作 → 全选表格
// 控件：btnSelectAllTables
namespace WordUpdater
{
    public partial class Connect
    {
        /// <summary>
        /// 选中当前文档中的全部表格（不含正文）。
        /// Word/WPS 的对象模型不支持"多段非连续选区"（Selection 只能是单段连续
        /// 区域），所以交给 TableSelectEngine 模拟人工的 Ctrl+框选：
        /// 逐个滚动到表格 → 取屏幕坐标 → 按住 Ctrl 模拟鼠标拖框。
        /// 选完后用户可以接着用功能区的表格工具 / 字体命令统一处理。
        /// </summary>
        public void OnSelectAllTables(object control)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                if (!EnsureApplication())
                {
                    UiDialog.Show(
                        "未检测到正在运行的 WPS 文字。\r\n请先打开文档，再点击「全选表格」。",
                        "全选表格", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                int total, selected;
                string err = TableSelectEngine.SelectAllTables(_app, out total, out selected);
                Log(string.Format("OnSelectAllTables -> total={0} selected={1} err={2}",
                    total, selected, err ?? "(null)"));
                if (err != null)
                {
                    UiDialog.Show(err, "全选表格",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (total == 0)
                {
                    UiDialog.Show("本文档的正文中没有表格。", "全选表格",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (selected < total)
                {
                    UiDialog.Show(string.Format(
                        "全文共 {0} 张表格，成功选中 {1} 张。\r\n"
                        + "有 {2} 张未能定位（可能被隐藏或在受保护区域），可手工 Ctrl 补选。",
                        total, selected, total - selected),
                        "全选表格", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                // 全部选中时不弹窗：选区本身就是反馈，弹窗反而打断后续操作
            }
            catch (Exception ex)
            {
                Log("OnSelectAllTables 异常：" + ex);
                try
                {
                    UiDialog.Show("发生未预期的错误：" + ex.Message, "全选表格",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
            }
            finally
            {
                _busy = false;
            }
        }

        public IPictureDisp GetImage_btnSelectAllTables(object control) { return GetButtonImageCore("btnSelectAllTables"); }
    }
}
