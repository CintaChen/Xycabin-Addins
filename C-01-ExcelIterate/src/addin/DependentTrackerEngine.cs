using System;
using System.Collections.Generic;

namespace Excel_Iterate
{
    // =====================================================================
    // 追踪从属引擎（碰 COM，遍历所有已打开工作簿反向匹配）
    //
    // 流程：
    //   1. 取目标单元格身份（工作簿名/工作表名/归一化地址）
    //   2. 遍历 app.Workbooks → 每个 wb.Worksheets → sheet.UsedRange
    //   3. 批量取 used.Formula（object[,] 二维数组，一次 COM 往返）内存遍历
    //   4. 对每个含公式的单元格，复用 FormulaRefParser.ParseAll 解析引用，
    //      逐项比对簿名+表名+地址（区域引用判断目标是否落在范围内）
    //   5. 进度回调 + 取消检查
    // =====================================================================

    /// <summary>一条从属结果：引用了目标单元格的那个单元格（引用者）。</summary>
    public class DependentItem
    {
        public int Index;
        public string WorkbookPath = "";    // 引用者工作簿完整路径
        public string WorkbookName = "";    // 引用者工作簿名
        public string SheetName = "";       // 引用者工作表名
        public string CellAddressA1 = "";   // 引用者单元格地址
        public string SourceFormula = "";   // 引用者公式原文
        public string MatchFragment = "";   // 命中的引用片段（引用目标的那段）
    }

    /// <summary>追踪从属引擎的输出。</summary>
    public class DependentResult
    {
        public TrackerStatus Status;
        public string ErrorMessage;
        public string TargetWorkbookName = "";   // 被追踪的目标单元格所在簿
        public string TargetSheetName = "";
        public string TargetCellAddress = "";
        public int ScannedCells;
        public int ScannedFormulas;
        public List<DependentItem> Items = new List<DependentItem>();
        public bool Cancelled;
    }

    /// <summary>扫描进度回调（后台线程调用，UI 层自行 BeginInvoke 更新）。</summary>
    public delegate void DependentProgress(int wbIndex, int wbCount, int sheetIndex,
        int sheetCount, string sheetName, int foundSoFar);

    public static class DependentTrackerEngine
    {
        /// <summary>追踪当前单元格的从属（反向找出引用它的全部单元格）。</summary>
        public static DependentResult TraceDependents(object application,
            DependentProgress progress, Func<bool> shouldCancel)
        {
            DependentResult r = new DependentResult();
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

                // 目标单元格身份
                r.TargetWorkbookName = (string)wb.Name;
                try { r.TargetSheetName = (string)app.ActiveSheet.Name; } catch { }
                string targetAddr = NormalizeAddr((string)cell.Address(false, false));
                r.TargetCellAddress = targetAddr;

                dynamic workbooks = app.Workbooks;
                int wbCount = (int)workbooks.Count;
                for (int w = 1; w <= wbCount; w++)
                {
                    if (shouldCancel != null && shouldCancel()) { r.Cancelled = true; break; }
                    dynamic curWb = GetItem(workbooks, w);
                    string wbName = (string)curWb.Name;
                    string wbPath = "";
                    try { wbPath = (string)curWb.FullName; } catch { }
                    dynamic sheets = curWb.Worksheets;
                    int sheetCount = (int)sheets.Count;
                    for (int s = 1; s <= sheetCount; s++)
                    {
                        if (shouldCancel != null && shouldCancel()) { r.Cancelled = true; break; }
                        dynamic sheet = GetItem(sheets, s);
                        string sheetName = (string)sheet.Name;
                        ScanSheet(sheet, wbName, wbPath, sheetName,
                            r.TargetWorkbookName, r.TargetSheetName, targetAddr, r);
                        if (progress != null)
                            progress(w, wbCount, s, sheetCount, sheetName, r.Items.Count);
                    }
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

        /// <summary>扫描单个工作表：批量取公式，逐格解析比对。</summary>
        private static void ScanSheet(dynamic sheet, string wbName, string wbPath, string sheetName,
            string targetWbName, string targetSheetName, string targetAddr, DependentResult r)
        {
            try
            {
                dynamic used = sheet.UsedRange;
                if (used == null) return;
                int firstRow = (int)used.Row;
                int firstCol = (int)used.Column;

                object formulaObj = used.Formula;
                if (formulaObj is object[,])
                {
                    object[,] arr = (object[,])formulaObj;
                    int rows = arr.GetLength(0);
                    int cols = arr.GetLength(1);
                    for (int rr = 1; rr <= rows; rr++)
                    {
                        for (int cc = 1; cc <= cols; cc++)
                        {
                            r.ScannedCells++;
                            object v = arr[rr, cc];
                            string f = v as string;
                            if (string.IsNullOrEmpty(f) || !f.StartsWith("=")) continue;
                            r.ScannedFormulas++;
                            int absRow = firstRow + rr - 1;
                            int absCol = firstCol + cc - 1;
                            string cellAddr = ColToLetters(absCol) + absRow.ToString();
                            MatchFormula(f, wbName, wbPath, sheetName, cellAddr,
                                targetWbName, targetSheetName, targetAddr, r);
                        }
                    }
                }
                else if (formulaObj is string)
                {
                    // 单格 UsedRange
                    r.ScannedCells++;
                    string f = (string)formulaObj;
                    if (!string.IsNullOrEmpty(f) && f.StartsWith("="))
                    {
                        r.ScannedFormulas++;
                        string cellAddr = ColToLetters(firstCol) + firstRow.ToString();
                        MatchFormula(f, wbName, wbPath, sheetName, cellAddr,
                            targetWbName, targetSheetName, targetAddr, r);
                    }
                }
            }
            catch { }
        }

        /// <summary>解析单个公式，检查是否引用了目标单元格；命中则加入结果。</summary>
        private static void MatchFormula(string formula, string wbName, string wbPath, string sheetName,
            string cellAddr, string targetWbName, string targetSheetName,
            string targetAddr, DependentResult r)
        {
            List<PrecedentItem> refs = FormulaRefParser.ParseAll(formula);
            for (int i = 0; i < refs.Count; i++)
            {
                PrecedentItem ri = refs[i];
                if (!MatchesTarget(ri, wbName, sheetName,
                    targetWbName, targetSheetName, targetAddr)) continue;

                DependentItem di = new DependentItem();
                di.Index = r.Items.Count + 1;
                di.WorkbookName = wbName;
                di.WorkbookPath = wbPath;
                di.SheetName = sheetName;
                di.CellAddressA1 = cellAddr;
                di.SourceFormula = formula;
                di.MatchFragment = ri.FormulaFragment;
                r.Items.Add(di);
            }
        }

        /// <summary>判断一个引用项是否指向目标单元格。</summary>
        private static bool MatchesTarget(PrecedentItem ri, string scanWbName, string scanSheetName,
            string targetWbName, string targetSheetName, string targetAddr)
        {
            // 地址匹配（区域引用判断目标是否落在范围内）
            bool addrMatch;
            if (ri.IsRange)
                addrMatch = InRange(targetAddr, ri.CellAddressA1, ri.RangeEndA1);
            else
                addrMatch = ri.CellAddressA1 == targetAddr;
            if (!addrMatch) return false;

            switch (ri.RefType)
            {
                case RefType.SameSheet:
                    // 同表引用指向扫描表自身
                    return scanWbName == targetWbName && scanSheetName == targetSheetName;
                case RefType.CrossSheet:
                    return scanWbName == targetWbName && ri.SheetName == targetSheetName;
                case RefType.CrossWorkbook:
                    // 跨簿引用的簿名可能含/不含扩展名，做前缀匹配
                    return MatchWbName(ri.WorkbookName, targetWbName)
                        && ri.SheetName == targetSheetName;
                default:
                    return false;
            }
        }

        private static bool MatchWbName(string refWbName, string targetWbName)
        {
            if (string.IsNullOrEmpty(refWbName) || string.IsNullOrEmpty(targetWbName)) return false;
            return refWbName == targetWbName || targetWbName.StartsWith(refWbName + ".");
        }

        private static string NormalizeAddr(string addr)
        {
            return (addr ?? "").Replace("$", "");
        }

        /// <summary>判断地址 addr 是否落在 [rangeStart, rangeEnd] 矩形范围内。</summary>
        private static bool InRange(string addr, string rangeStart, string rangeEnd)
        {
            try
            {
                int c1, r1, c2, r2, c, rw;
                ParseAddr(rangeStart, out c1, out r1);
                ParseAddr(rangeEnd, out c2, out r2);
                ParseAddr(addr, out c, out rw);
                int minC = Math.Min(c1, c2), maxC = Math.Max(c1, c2);
                int minR = Math.Min(r1, r2), maxR = Math.Max(r1, r2);
                return c >= minC && c <= maxC && rw >= minR && rw <= maxR;
            }
            catch { return false; }
        }

        /// <summary>把 A1 地址拆成 (列号, 行号)，均 1 起始。</summary>
        private static void ParseAddr(string addr, out int col, out int row)
        {
            int i = 0;
            while (i < addr.Length && char.IsLetter(addr[i])) i++;
            col = LettersToCol(addr.Substring(0, i));
            row = int.Parse(addr.Substring(i));
        }

        /// <summary>列号（1 起始）→ 列字母。</summary>
        private static string ColToLetters(int col)
        {
            string s = "";
            while (col > 0)
            {
                int m = (col - 1) % 26;
                s = (char)('A' + m) + s;
                col = (col - 1) / 26;
            }
            return s;
        }

        /// <summary>列字母 → 列号（1 起始）。</summary>
        private static int LettersToCol(string letters)
        {
            int col = 0;
            foreach (char ch in letters.ToUpperInvariant())
                col = col * 26 + (ch - 'A' + 1);
            return col;
        }

        /// <summary>COM 集合 1 起始索引访问。</summary>
        private static dynamic GetItem(dynamic collection, int index)
        {
            try { return collection.Item(index); }
            catch { return collection[index]; }
        }
    }
}
