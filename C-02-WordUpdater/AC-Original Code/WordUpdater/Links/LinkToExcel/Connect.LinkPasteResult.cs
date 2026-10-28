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
        private static string FormatLinkResult(LinkPasteResult r)
        {
            switch (r.Status)
            {
                case LinkPasteStatus.NoDocument:
                    return "当前没有打开的文档。";
                case LinkPasteStatus.NotInTable:
                    return "光标未位于表格内。\r\n\r\n请先点击到要链接的文档表格中，再点击「链接Excel」。";
                case LinkPasteStatus.TableNotUniform:
                    return "所选区域包含合并单元格、嵌套表格或不规则结构，无法与源区域逐格对应。\r\n\r\n请在同一张表格内选择由独立单元格组成的规则矩形区域后重试。";
                case LinkPasteStatus.TableTooLarge:
                    return string.Format("选中的区域共 {0} 行 × {1} 列 = {2} 个单元格，超出单次支持上限（{3} 个）。\r\n请缩小选中的范围后重试。",
                        r.TargetRows, r.TargetCols, r.TargetCells, LinkPasteEngine.MaxCells);
                case LinkPasteStatus.ClipboardInvalid:
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("剪贴板内容不是可链接的单元格区域。\r\n\r\n请先在 WPS 表格中复制要链接的单元格区域（可按住 Ctrl 多选多个区域），再点击「链接Excel」。");
                        if (!string.IsNullOrEmpty(r.ErrorMessage))
                        {
                            sb.Append("\r\n\r\n技术详情：");
                            sb.Append(r.ErrorMessage);
                        }
                        return sb.ToString();
                    }
                case LinkPasteStatus.MultiArea:
                    return "复制的是多个不连续的单元格区域，但剪贴板没能还原它们的明细位置。\r\n\r\n请保持 WPS 表格处于打开状态、且当前选区仍是刚复制的那几个格子，然后点击「链接Excel」重试。";
                case LinkPasteStatus.SizeMismatch:
                    {
                        bool overflow = r.SourceCells > r.TargetCells;
                        string head = overflow
                            ? "源数据的格数多于选中的区域（区域溢出）。"
                            : "源数据的格数少于选中的区域（区域未填满）。";
                        string tail = overflow
                            ? "请扩大选中的区域，或改复制更小的区域后重试。"
                            : "请缩小选中的区域，或补齐复制的格子后重试。";
                        return string.Format(
                            head + "\r\n\r\n复制的源：{0} 个单元格\r\n选中的区域：{1} 行 × {2} 列 = {3} 个单元格"
                            + "\r\n\r\n链接无效，文档中未作任何修改。\r\n" + tail,
                            r.SourceCells, r.TargetRows, r.TargetCols, r.TargetCells);
                    }
                case LinkPasteStatus.SelectionEmpty:
                    return "没有选中任何表格单元格。\r\n\r\n请先在表格里拖选出要链接的矩形区域"
                        + "（要链接整张表格，就把整张表都选上），再点击「链接Excel」。";
                case LinkPasteStatus.SelectionNotRect:
                    return "所选区域包含合并单元格、嵌套表格或不规则结构，无法逐格对应。\r\n\r\n" + (r.ErrorMessage ?? "")
                        + "\r\n\r\n请重新选择由独立单元格组成、列边界对齐的连续矩形区域后重试。";
                case LinkPasteStatus.SelectionOutOfTable:
                    return "选区超出了表格范围。\r\n\r\n" + (r.ErrorMessage ?? "")
                        + "\r\n\r\n请只在同一张表格内部拖选矩形区域（不要拖到表格外的正文，也不要跨表）。";
                case LinkPasteStatus.SourceMissing:
                    return "源工作簿文件不存在或已被移动：\r\n\r\n" + (r.SourceDesc ?? "") +
                        "\r\n\r\n请确认源文件位置后重试。";
                case LinkPasteStatus.Error:
                    return "链接粘贴未能完成" +
                        (r.RolledBack ? "（已自动撤销已写入的部分，未产生残缺链接）。\r\n" : "。\r\n") +
                        "错误信息：" + (r.ErrorMessage ?? "未知");
                case LinkPasteStatus.Completed:
                    {
                        string tpl = RibbonConfig.GetResultText("OnLinkToExcel");
                        if (tpl != null)
                        {
                            // 用户在 配置器脚本 里自定义了完成文案
                            StringBuilder t = new StringBuilder();
                            t.Append(FillPasteTemplate(tpl, r));
                            if (r.Failures.Count > 0)
                            {
                                // 刷新失败明细是安全信息，自定义模板也照附（配置里说明过）
                                t.AppendLine();
                                t.AppendLine();
                                t.AppendLine("刷新失败明细：");
                                int maxF = Math.Min(r.Failures.Count, 5);
                                for (int i = 0; i < maxF; i++)
                                    t.Append("  " + r.Failures[i] + "\r\n");
                                if (r.Failures.Count > maxF)
                                    t.AppendFormat("  …… 其余 {0} 条略。", r.Failures.Count - maxF);
                            }
                            if (r.CleanedHiddenHosts > 0 &&
                                tpl.IndexOf("{cleaned}", StringComparison.Ordinal) < 0)
                            {
                                t.AppendLine();
                                t.AppendFormat("（已自动释放 {0} 个刷新时拉起的后台 WPS 表格进程，源文件占用已解除。）",
                                    r.CleanedHiddenHosts);
                            }
                            return t.ToString();
                        }

                        StringBuilder sb = new StringBuilder();
                        sb.AppendFormat("链接粘贴完成：已在选中的 {0} 行 × {1} 列里为 {2} 个单元格建立链接",
                            r.TargetRows, r.TargetCols, r.CellsLinked);
                        if (r.Refreshed >= r.CellsLinked && r.Failures.Count == 0)
                        {
                            sb.Append("，并已全部刷新。");
                        }
                        else
                        {
                            sb.AppendFormat("，其中 {0} 个已刷新。", r.Refreshed);
                            if (r.Failures.Count > 0)
                            {
                                sb.Append("\r\n其余刷新失败（链接本身有效，源数据可用时可选中表格点「更新链接」重试）。");
                            }
                        }
                        sb.AppendLine();
                        sb.AppendLine();
                        sb.Append("源：");
                        sb.Append(string.IsNullOrEmpty(r.SourceDesc) ? "(未知)" : r.SourceDesc);
                        sb.AppendLine();
                        sb.Append("目标：");
                        if (r.TargetRows == 1 && r.TargetCols == 1)
                            sb.Append("第 " + r.TargetRow1 + " 行第 " + r.TargetCol1 + " 列");
                        else
                            sb.Append("第 " + r.TargetRow1 + "~" + (r.TargetRow1 + r.TargetRows - 1)
                                + " 行 × 第 " + r.TargetCol1 + "~" + (r.TargetCol1 + r.TargetCols - 1) + " 列");
                        sb.AppendLine();
                        sb.AppendLine();
                        sb.Append("提示：源数据变动后，选中该表格并点击「更新链接」即可一键同步。");
                        if (r.Failures.Count > 0)
                        {
                            sb.AppendLine();
                            sb.AppendLine();
                            sb.AppendLine("刷新失败明细：");
                            int max = Math.Min(r.Failures.Count, 5);
                            for (int i = 0; i < max; i++)
                                sb.Append("  " + r.Failures[i] + "\r\n");
                            if (r.Failures.Count > max)
                                sb.AppendFormat("  …… 其余 {0} 条略。", r.Failures.Count - max);
                        }
                        if (r.CleanedHiddenHosts > 0)
                        {
                            sb.AppendLine();
                            sb.AppendFormat("（已自动释放 {0} 个刷新时拉起的后台 WPS 表格进程，源文件占用已解除。）",
                                r.CleanedHiddenHosts);
                        }
                        return sb.ToString();
                    }
                default:
                    return "未知状态。";
            }
        }

        /// <summary>把链接Excel的自定义模板里的占位符换成实际数字/文本。</summary>
        private static string FillPasteTemplate(string tpl, LinkPasteResult r)
        {
            return tpl
                .Replace("{rows}", r.TargetRows.ToString())
                .Replace("{cols}", r.TargetCols.ToString())
                .Replace("{cells}", r.CellsLinked.ToString())
                .Replace("{refreshed}", r.Refreshed.ToString())
                .Replace("{failures}", r.Failures.Count.ToString())
                .Replace("{source}", string.IsNullOrEmpty(r.SourceDesc) ? "(未知)" : r.SourceDesc)
                .Replace("{cleaned}", r.CleanedHiddenHosts.ToString());
        }
    }
}
