using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

// 功能：WordUpdater → 链接 → 链接Excel
// 控件：btnLinkToExcel
namespace WordUpdater
{
    public partial class Connect
    {
        public void OnLinkToExcel(object control)
        {
            if (_busy) return;
            _busy = true;
            try
            {
                if (!EnsureApplication())
                {
                    UiDialog.Show(
                        "未检测到正在运行的 WPS 文字。\r\n请先打开 WPS 文字文档，再点击「链接Excel」。",
                        "链接Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                LinkPasteResult result = LinkPasteEngine.Execute(_app);
                Log(string.Format("OnLinkToExcel -> {0} linked={1} refreshed={2} rolledBack={3} src={4}",
                    result.Status, result.CellsLinked, result.Refreshed, result.RolledBack,
                    result.SourceDesc));
                MessageBoxIcon icon = result.Status == LinkPasteStatus.Error
                    ? MessageBoxIcon.Error
                    : (result.Status == LinkPasteStatus.Completed
                        ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
                UiDialog.Show(FormatLinkResult(result), "链接Excel",
                    MessageBoxButtons.OK, icon);
            }
            catch (Exception ex)
            {
                Log("OnLinkToExcel 异常：" + ex);
                try
                {
                    UiDialog.Show("发生未预期的错误：" + ex.Message, "链接Excel",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
            }
            finally
            {
                _busy = false;
            }
        }

        public IPictureDisp GetImage_btnLinkToExcel(object control) { return GetButtonImageCore("btnLinkToExcel"); }
    }
}
