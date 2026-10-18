using System;
using System.Collections.Generic;
using System.IO;

namespace Excel_Iterate
{
    // =====================================================================
    // 追踪引用引擎（碰 COM，通过 WPS 表格对象模型取公式再调 FormulaRefParser）
    //
    // 流程：
    //   1. COM 取 ActiveWorkbook / ActiveCell / cell.Formula
    //   2. 调 FormulaRefParser.ParseAll(formula) 纯解析
    //   3. 补全 SameSheet（填当前簿/表）/ CrossWorkbook（查已开簿确认路径）
    //   4. 输出 PrecedentResult
    // =====================================================================

    public static class PrecedentTrackerEngine
    {
        /// <summary>追踪当前单元格的引用项（追踪引用功能核心）。</summary>
        public static PrecedentResult TracePrecedents(object application)
        {
            PrecedentResult r = new PrecedentResult();
            if (application == null)
            {
                r.Status = TrackerStatus.Error;
                r.ErrorMessage = "宿主为空";
                return r;
            }
            try
            {
                dynamic app = application;
                dynamic wb = app.ActiveWorkbook;
                if (wb == null) { r.Status = TrackerStatus.NoWorkbook; return r; }
                dynamic cell = app.ActiveCell;
                if (cell == null) { r.Status = TrackerStatus.NoActiveCell; return r; }

                r.WorkbookName = (string)wb.Name;
                try { r.WorkbookPath = (string)wb.FullName; } catch { }
                try { r.SheetName = (string)app.ActiveSheet.Name; } catch { }
                try { r.CellAddress = (string)cell.Address(false, false); } catch { }

                string formula = "";
                try { formula = (string)cell.Formula; } catch { }
                if (string.IsNullOrEmpty(formula) || !formula.StartsWith("="))
                {
                    r.Status = TrackerStatus.NoFormula;
                    return r;
                }
                r.Formula = formula;

                // 纯解析
                List<PrecedentItem> items = FormulaRefParser.ParseAll(formula);

                // 补全 SameSheet / CrossSheet / CrossWorkbook 信息（需要 COM）
                string curWbName = r.WorkbookName;
                string curWbPath = "";
                try { curWbPath = (string)wb.FullName; } catch { }
                string curSheet = r.SheetName;

                for (int i = 0; i < items.Count; i++)
                {
                    PrecedentItem item = items[i];
                    if (item.RefType == RefType.SameSheet)
                    {
                        item.WorkbookPath = curWbPath;
                        item.WorkbookName = curWbName;
                        item.SheetName = curSheet;
                    }
                    else if (item.RefType == RefType.CrossSheet)
                    {
                        item.WorkbookPath = curWbPath;
                        item.WorkbookName = curWbName;
                    }
                    else if (item.RefType == RefType.CrossWorkbook)
                    {
                        ResolveCrossWorkbook(app, item);
                    }
                    r.Items.Add(item);
                }
                r.Status = TrackerStatus.Completed;
                return r;
            }
            catch (Exception ex)
            {
                r.Status = TrackerStatus.Error;
                r.ErrorMessage = ex.Message;
                return r;
            }
        }

        /// <summary>
        /// 查已打开工作簿确认跨簿引用的完整路径。
        /// 早期 Excel 语法 [工作簿1] 的工作簿名可能无扩展名，
        /// 先按 Name == workbookName 或 Name.StartsWith(workbookName + ".") 匹配已开簿；
        /// 未找到则用路径 + 工作簿名 + .xlsx 推测（UI 会标注"路径为推断"）。
        /// </summary>
        private static void ResolveCrossWorkbook(dynamic app, PrecedentItem item)
        {
            try
            {
                dynamic wbs = app.Workbooks;
                int count = (int)wbs.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic w = GetItem(wbs, i);
                    string name = (string)w.Name;
                    // 精确匹配 或 前缀匹配（工作簿名可能无扩展名）
                    if (name == item.WorkbookName
                        || name.StartsWith(item.WorkbookName + "."))
                    {
                        try { item.WorkbookPath = (string)w.FullName; } catch { }
                        item.WorkbookName = name;
                        return;
                    }
                }
            }
            catch { }
            // 未找到已打开工作簿：用路径 + 工作簿名推测完整路径
            if (!string.IsNullOrEmpty(item.WorkbookPath))
            {
                string wbName = item.WorkbookName;
                // 工作簿名已含扩展名（含 .）则直接用，否则补 .xlsx
                if (wbName.IndexOf('.') < 0) wbName += ".xlsx";
                item.WorkbookPath = Path.Combine(item.WorkbookPath, wbName);
            }
        }

        /// <summary>COM 集合 1 起始索引访问（dynamic 下索引器发出 Item 方法名）。</summary>
        private static dynamic GetItem(dynamic collection, int index)
        {
            try { return collection.Item(index); }
            catch { return collection[index]; }
        }

        // ==================================================================
        //                          跳转 + 标黄（ET 用 Interior.Color）
        // ==================================================================

        // 黄色 OLE 颜色 65535 (0xFFFF)
        private const int FlashYellow = 65535;

        // xlNone = -4142：Interior.ColorIndex 为负 = 无显式填充（"无填充"）。
        // 注意 Interior.Color 对无填充格返回白色 16777215，不能拿它当还原值，
        // 否则会把"无填充"误还原成"白色填充"。这里用 -4142 当哨兵：还原时走 ColorIndex 清填充。
        private const int NoFill = -4142;

        /// <summary>跳转到引用项指向的单元格并标黄（薄包装，转调 JumpToCellEx）。</summary>
        public static string JumpToPrecedentEx(object application, PrecedentItem item,
            out object flashRange, out int prevColor)
        {
            flashRange = null;
            prevColor = 0;
            if (item == null) return "参数为空。";
            return JumpToCellEx(application, item.WorkbookName, item.WorkbookPath,
                item.SheetName, item.CellAddressA1, item.RefType == RefType.CrossWorkbook,
                out flashRange, out prevColor);
        }

        /// <summary>
        /// 通用跳转：按工作簿名/表名/地址定位单元格并标黄。
        /// 目标信息与活动单元格无关，支持连续跳转；跨簿未打开时自动打开。
        /// 标黄先读原值，还原按原值（非清零）；失败静默不影响跳转。
        /// </summary>
        public static string JumpToCellEx(object application,
            string workbookName, string workbookPath, string sheetName, string cellAddress,
            bool isCrossWorkbook, out object flashRange, out int prevColor)
        {
            flashRange = null;
            prevColor = 0;
            if (application == null) return "参数为空。";
            try
            {
                dynamic app = application;

                // 定位目标工作簿：统一按名称查找，不依赖 ActiveWorkbook
                dynamic wb = FindOpenWorkbook(app, workbookName, workbookPath);
                if (wb == null && isCrossWorkbook && !string.IsNullOrEmpty(workbookPath))
                {
                    // 跨簿未打开：自动打开
                    if (!File.Exists(workbookPath))
                    {
                        return "目标工作簿不存在：\r\n" + workbookPath
                            + "\r\n\r\n文件可能已被移动、重命名或删除。";
                    }
                    try
                    {
                        wb = app.Workbooks.Open(workbookPath);
                    }
                    catch (Exception openEx)
                    {
                        return "打开目标工作簿失败：" + openEx.Message + "\r\n路径：" + workbookPath;
                    }
                }
                if (wb == null)
                {
                    // 兜底：同表/跨表按名称找不到时退回活动工作簿
                    try { wb = app.ActiveWorkbook; } catch { }
                }
                if (wb == null) return "目标工作簿未找到：\r\n" + workbookName;

                dynamic ws = FindSheet(wb, sheetName);
                if (ws == null) return "工作表「" + sheetName + "」未找到。";
                dynamic targetRange = ws.Range[cellAddress];

                // 激活 + 选定 + 标黄
                try { targetRange.Worksheet.Activate(); } catch { }
                try { targetRange.Select(); } catch { }
                try
                {
                    dynamic interior = targetRange.Interior;
                    prevColor = (int)interior.Color;     // 先读原值（有效 RGB）
                    int colorIndex = NoFill;
                    try { colorIndex = (int)interior.ColorIndex; } catch { }
                    interior.Color = FlashYellow;        // 写黄
                    if (colorIndex < 0) prevColor = NoFill;   // 原为「无填充」：记哨兵，还原时清填充而非涂白
                    flashRange = targetRange;
                }
                catch { flashRange = null; prevColor = 0; }

                // 焦点交还宿主
                try { app.ActiveWindow.Activate(); } catch { try { app.Activate(); } catch { } }
                return null;
            }
            catch (Exception ex)
            {
                return "跳转失败：" + ex.Message;
            }
        }

        /// <summary>
        /// 还原跳转标黄的区域（按原值还原，静默失败）。
        /// prevColor == NoFill(-4142) 表示原为「无填充」→ 改回 ColorIndex 清填充；
        /// 否则按原 RGB 精确还原。
        /// </summary>
        public static void ClearFlash(object range, int prevColor)
        {
            if (range == null) return;
            try
            {
                dynamic r = range;
                if (prevColor == NoFill) r.Interior.ColorIndex = NoFill;
                else r.Interior.Color = prevColor;
            }
            catch { }
        }

        /// <summary>按名称在 Workbook.Worksheets 中查找工作表；找不到返回 null。</summary>
        internal static dynamic FindSheet(dynamic wb, string sheetName)
        {
            try
            {
                dynamic sheets = wb.Worksheets;
                for (int i = 1; i <= (int)sheets.Count; i++)
                {
                    dynamic ws = GetItem(sheets, i);
                    if ((string)ws.Name == sheetName) return ws;
                }
            }
            catch { }
            return null;
        }

        /// <summary>在已打开工作簿中查找；找不到返回 null。</summary>
        internal static dynamic FindOpenWorkbook(dynamic app, string wbName, string wbPath)
        {
            try
            {
                dynamic wbs = app.Workbooks;
                for (int i = 1; i <= (int)wbs.Count; i++)
                {
                    dynamic w = GetItem(wbs, i);
                    string n = (string)w.Name;
                    if (n == wbName || n.StartsWith(wbName + ".")) return w;
                }
            }
            catch { }
            return null;
        }
    }
}
