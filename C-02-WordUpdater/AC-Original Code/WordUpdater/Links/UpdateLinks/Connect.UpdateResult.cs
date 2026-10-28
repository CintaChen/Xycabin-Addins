using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace WordUpdater
{
    public partial class Connect
    {
        /// <summary>把更新链接的自定义模板里的占位符换成实际数字。</summary>
        private static string FillUpdateTemplate(string tpl, RunResult r)
        {
            return tpl
                .Replace("{found}", r.Found.ToString())
                .Replace("{success}", r.Success.ToString())
                .Replace("{failed}", r.Failed.ToString())
                .Replace("{changed}", r.Changes.Count.ToString())
                .Replace("{cleaned}", r.CleanedHiddenHosts.ToString());
        }

        private static string FormatResult(RunResult r)
        {
            switch (r.Status)
            {
                case RunStatus.NoDocument:
                    return "当前没有打开的文档。";
                case RunStatus.NoSelection:
                    return "请先选中包含链接的内容范围（按住鼠标拖选，或按 Ctrl+A 全选），再点击「更新链接」。";
                case RunStatus.NoLinks:
                    return "所选内容中未发现任何链接。\r\n\r\n提示：从 Excel 复制单元格后，在 WPS 文字中「粘贴」→「选择性粘贴」→「粘贴链接」，即可创建可更新的链接。";
                case RunStatus.Error:
                    return "更新过程中发生错误：\r\n" + (r.ErrorMessage ?? "未知错误");
                case RunStatus.Completed:
                    {
                        StringBuilder sb = new StringBuilder();
                        string tpl = RibbonConfig.GetResultText("OnUpdateLinks");
                        if (tpl != null)
                        {
                            // 用户在 配置器脚本 里自定义了完成文案
                            sb.Append(FillUpdateTemplate(tpl, r));
                        }
                        else
                        {
                            sb.AppendFormat("更新完成：共发现 {0} 个链接，成功 {1} 个，数值变动 {2} 个",
                                r.Found, r.Success, r.Changes.Count);
                            if (r.Failed > 0)
                            {
                                sb.AppendFormat("，失败 {0} 个。", r.Failed);
                            }
                            else
                            {
                                sb.Append("。");
                            }
                        }

                        // 变动明细：哪些链接的数值变了（这些位置已标黄）
                        if (r.Changes.Count > 0)
                        {
                            sb.AppendLine();
                            sb.AppendLine();
                            sb.AppendLine("变动明细（已标黄，关闭本窗口后 2.5 秒自动还原）：");
                            int max = Math.Min(r.Changes.Count, 10);
                            for (int i = 0; i < max; i++)
                            {
                                LinkChangeInfo c = r.Changes[i];
                                string oldText = c.OldText.Length == 0 ? "（空）" : c.OldText;
                                string newText = c.NewText.Length == 0 ? "（空）" : c.NewText;
                                sb.AppendFormat("  {0}. {1}：{2} → {3}\r\n",
                                    c.Index, c.Where, oldText, newText);
                            }
                            if (r.Changes.Count > max)
                                sb.AppendFormat("  …… 其余 {0} 条变动略。", r.Changes.Count - max);
                        }

                        if (r.Failed > 0)
                        {
                            // 失败明细是安全信息，自定义模板也照附（配置里说明过）
                            sb.AppendLine();
                            sb.AppendLine();
                            sb.AppendLine("失败明细：");
                            int max = Math.Min(r.Failures.Count, 10);
                            for (int i = 0; i < max; i++)
                            {
                                LinkFailureInfo f = r.Failures[i];
                                string src = string.IsNullOrEmpty(f.Source) ? "(未知)" : f.Source;
                                sb.AppendFormat("  {0}. [{1}] 源：{2}\r\n     原因：{3}\r\n",
                                    f.Index, f.Kind, src, f.Reason);
                            }
                            if (r.Failures.Count > max)
                                sb.AppendFormat("  …… 其余 {0} 条失败信息略。", r.Failures.Count - max);
                        }
                        // 更新成功但没有可比对文本的：链接图片/对象（或读不到文本）
                        if (r.NoTextCount > 0)
                        {
                            sb.AppendLine();
                            sb.AppendFormat("其中 {0} 个为图片/对象，无文本可比。", r.NoTextCount);
                        }
                        if (r.CleanedHiddenHosts > 0 &&
                            (tpl == null || tpl.IndexOf("{cleaned}", StringComparison.Ordinal) < 0))
                        {
                            sb.AppendLine();
                            sb.AppendFormat("（已自动释放 {0} 个更新时拉起的后台 WPS 表格进程，源文件占用已解除。）",
                                r.CleanedHiddenHosts);
                        }
                        return sb.ToString();
                    }
                default:
                    return "未知状态。";
            }
        }
    }
}
