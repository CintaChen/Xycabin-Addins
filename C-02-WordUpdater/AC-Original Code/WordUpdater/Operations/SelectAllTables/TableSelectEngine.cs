using System;

namespace WordUpdater
{
    // =====================================================================
    // 「全选表格」引擎 —— 可编辑区域（Editors）技巧
    //
    //   Word/WPS 的 COM 对象模型里 Selection 只能是【单段连续区域】，
    //   没有多段选区 API；模拟鼠标 Ctrl+框选的方案实机不可靠（已废弃）。
    //
    //   但 Word 界面层本来就支持多段选区 ——「限制编辑」里的【可编辑区域】
    //   可以被一次性全部选中。于是借道（与经典 VBA 宏 SelectAllTables 同款）：
    //     1) 给每张表格的 Range 临时挂一个"人人可编辑"标记（Editors.Add）
    //     2) Document.SelectAllEditableRanges 把所有带标记的区域一次选中
    //        —— 真·多段选区，与手工 Ctrl 逐个框选的结果完全等价
    //     3) 立刻删掉这些临时标记（DeleteAllEditableRanges）
    //        关键：删标记【不会】撤销已生成的选区，选区原样保留
    //
    //   全程纯对象模型、瞬间完成、不滚屏不动鼠标；
    //   只做"选中"不改内容（Editors 标记随加随删，不落文档）。
    //
    //   已知边界：
    //     · 只处理正文顶层表格（页眉页脚/文本框/嵌套表格不在内）；
    //     · 文档处于「表单填写」保护时此技巧不可用（会明确提示）；
    //     · Editors.Add 失败的个别表格（受保护区域等）跳过并计数。
    // =====================================================================
    internal static class TableSelectEngine
    {
        // Word 常量：wdEditorEveryone = -1；wdAllowOnlyFormFields = 2
        private const int WdEditorEveryone = -1;
        private const int WdAllowOnlyFormFields = 2;

        /// <summary>
        /// 选中当前文档中的全部表格。返回 null = 流程走完（total==0 表示没有表格）；
        /// 返回字符串 = 失败原因（交 UI 弹给用户）。selected&lt;total 表示部分未选中。
        /// </summary>
        public static string SelectAllTables(object application, out int total, out int selected)
        {
            total = 0;
            selected = 0;
            if (application == null) return "未获取到 WPS 文字。";

            dynamic app = application;
            bool? screenUpdating = null;
            try
            {
                dynamic doc = app.ActiveDocument;
                total = (int)doc.Tables.Count;
                if (total == 0) return null;

                // 表单保护下 Editors 技巧不可用，提前给出明确提示
                int protection = -1;
                try { protection = (int)doc.ProtectionType; } catch { }
                if (protection == WdAllowOnlyFormFields)
                    return "文档处于「表单填写」保护状态，无法批量选中表格。\r\n"
                        + "请先在「审阅 → 限制编辑」中停止保护，或手工按住 Ctrl 逐个框选。";

                try { screenUpdating = (bool)app.ScreenUpdating; } catch { }
                try { app.ScreenUpdating = false; } catch { }

                // 先清一遍历史残留的"人人可编辑"标记（上次异常退出可能留下）
                try { doc.DeleteAllEditableRanges(WdEditorEveryone); } catch { }

                // 逐表挂临时标记：Range.Editors.Add(wdEditorEveryone)
                for (int i = 1; i <= total; i++)
                {
                    try
                    {
                        dynamic tbl = doc.Tables.Item(i);
                        tbl.Range.Editors.Add(WdEditorEveryone);
                        selected++;
                    }
                    catch { /* 个别表格挂不上（受保护区域等），跳过并由计数反映 */ }
                }
                if (selected == 0)
                    return "未能选中任何表格（共 " + total + " 张）。\r\n"
                        + "文档可能处于受保护/只读状态，请解除后重试，"
                        + "或手工按住 Ctrl 逐个框选。";

                // 一次性选中全部带标记区域 —— 真·多段选区在此成形
                doc.SelectAllEditableRanges(WdEditorEveryone);

                // 删掉临时标记；选区保留（这正是技巧的要点）
                try { doc.DeleteAllEditableRanges(WdEditorEveryone); } catch { }
                return null;
            }
            catch (Exception ex)
            {
                // 异常路径上也要尽力清标记，免得在文档里留下"可编辑区域"残留
                try { app.ActiveDocument.DeleteAllEditableRanges(WdEditorEveryone); } catch { }
                return "全选表格失败：" + ex.Message;
            }
            finally
            {
                if (screenUpdating.HasValue)
                {
                    try { app.ScreenUpdating = screenUpdating.Value; } catch { }
                }
            }
        }
    }
}
