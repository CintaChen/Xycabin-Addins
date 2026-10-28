using System;
using System.Collections.Generic;
using System.IO;

namespace Excel_Iterate
{
    // =====================================================================
    // 错误来源追溯引擎（「显示错误来源」核心）
    //
    // 痛点：A2 显示 #REF! 但公式 =A1+A2+A3 不含错误字面量，沿链
    //       A3=B3+B4、B4=D12+'#REF!'!I22，根因是 B4。
    // 判据三层（对每格依次判定）：
    //   L0 文本层：公式剔除双引号串后含错误字面量（#REF! 等）⇒ 错误源
    //   L1 值层：  单元格当前值解码为错误 ⇒ 错误节点；非错误 ⇒ 剪枝（不再回溯上游）
    //   L2 引用层：错误节点中，上游有错误 ⇒ 传播节点（继续回溯）；
    //              上游全正常 ⇒ 错误源（自算错误，如 =1/0）；无公式 ⇒ 常量（剪枝）
    //
    // 分层设计：
    //   TraceCore 纯核：只依赖 CellReader 委托读格，可注入假数据离线自测。
    //   TraceFrom + ComCellReader：碰 COM 的入口（读 Value2/Value/Text、自动开簿）。
    // =====================================================================

    /// <summary>错误来源判定原因。</summary>
    public enum ErrorSourceReason
    {
        /// <summary>公式文本含错误字面量（如 '#REF!'!I22 或 ='#REF!'+1）。</summary>
        LiteralInFormula,
        /// <summary>自算错误：本格算错但全部上游正常（如 =1/0）。</summary>
        SelfComputed
    }

    /// <summary>一个错误源单元格（回溯终点）。</summary>
    public class ErrorSourceItem
    {
        public string WorkbookPath = "";
        public string WorkbookName = "";
        public string SheetName = "";
        public string CellAddressA1 = "";
        public string Formula = "";          // 该格公式原文（可能为空）
        public string ErrorText = "";        // 解码出的错误文本（#REF! 等）
        public ErrorSourceReason Reason;
        public int Depth;                    // 距起点的回溯层数（0=起点自身）
        public string PathText = "";         // 传播路径，形如 A2 ← A3 ← B4
        public bool IsCrossWorkbook;         // 供跳转用
    }

    /// <summary>遍历上限与开关（默认值即 UI 使用值）。</summary>
    public class ErrorTraceOptions
    {
        public int MaxDepth = 12;            // 最大回溯层数
        public int MaxNodes = 300;           // 最大访问节点数（防爆炸）
        public int MaxSources = 50;          // 最多收集的错误源个数
        public int MaxAutoOpen = 20;         // 自动打开外部工作簿的数量上限
        public bool AutoOpenWorkbook = true; // 跨簿未打开时自动 Open
        public TimeSpan TimeLimit = TimeSpan.FromSeconds(8); // 总耗时上限
    }

    /// <summary>错误来源追溯结果。</summary>
    public class ErrorTraceResult
    {
        public bool Success;
        public bool Cancelled;               // 用户取消
        public bool TruncatedByLimit;        // 触到任一上限（深度/节点/源数/耗时/开簿数）
        public string ErrorMessage = "";
        public string LimitNote = "";        // 触限说明（供 UI 提示"结果可能不完整"）
        public int NodesVisited;
        public int AutoOpened;               // 本次自动打开的工作簿数
        public List<ErrorSourceItem> Sources = new List<ErrorSourceItem>();
    }

    /// <summary>一次单元格读取的结果（CellReader 的返回）。</summary>
    public class CellReadResult
    {
        public bool Found;                   // 工作簿/表/格定位成功
        public bool AutoOpened;              // 本次读取是否自动打开了外部工作簿
        public string Formula = "";          // 公式原文（无公式为空）
        public bool HasErrorValue;           // 当前值是错误值
        public string ErrorText = "";        // 错误文本（#REF! 等）；HasErrorValue=false 时为空
        public string FailNote = "";         // 定位失败原因（供 UI 提示）
    }

    /// <summary>读一格的委托：给定簿路径/簿名/表名/A1 地址，返回公式与错误值信息。</summary>
    public delegate CellReadResult CellReader(
        string workbookPath, string workbookName, string sheetName, string cellAddressA1);

    /// <summary>错误来源追溯入口（UI 调 TraceFrom，探针调 TraceCore）。</summary>
    public static class ErrorSourceEngine
    {
        // ---- Excel 错误值解码（CVErr 常量 ↔ 文本）----

        // CVErr 错误码（单元格 Value2 返回 int 错误码，或带 0x800A0000 高位）
        const int ErrNull = 2000;   // #NULL!
        const int ErrDiv0 = 2007;   // #DIV/0!
        const int ErrValue = 2015;  // #VALUE!
        const int ErrRef = 2023;    // #REF!
        const int ErrName = 2029;   // #NAME?
        const int ErrNum = 2036;    // #NUM!
        const int ErrNA = 2042;     // #N/A

        /// <summary>把 COM 返回值（Value2/Value/Text 形态之一）解码为错误文本；非错误返回 null。</summary>
        public static string TryReadError(object value)
        {
            if (value == null) return null;
            // Text 形态：直接是字符串 "#REF!"
            string s = value as string;
            if (s != null)
            {
                for (int i = 0; i < ErrorLiteralsForText.Length; i++)
                    if (string.Equals(s, ErrorLiteralsForText[i], StringComparison.OrdinalIgnoreCase))
                        return ErrorLiteralsForText[i];
                return null;
            }
            // int 形态：错误码本体或带 HRESULT 高位 0x800A0000
            int code;
            try { code = Convert.ToInt32(value); }
            catch { return null; }
            int low = code & 0xFFFF;   // 0x800A07FA → 0x07FA = 2042
            return ErrorCodeToText(low);
        }

        static readonly string[] ErrorLiteralsForText =
            { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A" };

        /// <summary>CVErr 错误码 → 错误文本；未知码返回 null。</summary>
        public static string ErrorCodeToText(int code)
        {
            switch (code)
            {
                case ErrNull: return "#NULL!";
                case ErrDiv0: return "#DIV/0!";
                case ErrValue: return "#VALUE!";
                case ErrRef: return "#REF!";
                case ErrName: return "#NAME?";
                case ErrNum: return "#NUM!";
                case ErrNA: return "#N/A";
            }
            return null;
        }

        // ---- TraceCore：纯核 BFS（不碰 COM，CellReader 可注入假数据）----

        class QNode
        {
            public string WbPath, WbName, Sheet, Addr;
            public int Depth;
            public string PathText;          // 从起点到本格的传播路径（A2 ← A3 ← B4）
            public bool IsCrossWorkbook;
        }

        /// <summary>
        /// 纯核回溯：从若干起点格（rootAddrs，同属 rootSheet/rootWb）沿 Precedents 向上，
        /// 收集全部错误源。reader 负责读格（可注入假数据离线测）。
        /// shouldCancel 返回 true 时尽快退出（Cancelled=true）。
        /// </summary>
        public static ErrorTraceResult TraceCore(
            string rootWbPath, string rootWbName, string rootSheet, IList<string> rootAddrs,
            CellReader reader, ErrorTraceOptions options, Func<bool> shouldCancel)
        {
            ErrorTraceResult r = new ErrorTraceResult();
            if (reader == null) { r.ErrorMessage = "读取器为空"; return r; }
            if (options == null) options = new ErrorTraceOptions();
            if (rootAddrs == null || rootAddrs.Count == 0) { r.ErrorMessage = "起点为空"; return r; }

            DateTime start = DateTime.Now;
            Queue<QNode> queue = new Queue<QNode>();
            HashSet<string> visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < rootAddrs.Count; i++)
            {
                string addr = rootAddrs[i];
                if (string.IsNullOrEmpty(addr)) continue;
                QNode n = new QNode();
                n.WbPath = rootWbPath ?? ""; n.WbName = rootWbName ?? "";
                n.Sheet = rootSheet ?? ""; n.Addr = addr;
                n.Depth = 0; n.PathText = addr;
                string key = NodeKey(n);
                if (visited.Add(key)) queue.Enqueue(n);
            }

            while (queue.Count > 0)
            {
                if (shouldCancel != null && shouldCancel()) { r.Cancelled = true; break; }
                if (DateTime.Now - start > options.TimeLimit)
                {
                    r.TruncatedByLimit = true;
                    r.LimitNote = "超过 " + (int)options.TimeLimit.TotalSeconds + " 秒时限";
                    break;
                }
                if (r.NodesVisited >= options.MaxNodes)
                {
                    r.TruncatedByLimit = true;
                    r.LimitNote = "达到节点上限 " + options.MaxNodes;
                    break;
                }

                QNode cur = queue.Dequeue();
                r.NodesVisited++;

                CellReadResult cr;
                try { cr = reader(cur.WbPath, cur.WbName, cur.Sheet, cur.Addr); }
                catch (Exception ex)
                {
                    r.TruncatedByLimit = true;
                    r.LimitNote = "读取 " + cur.Addr + " 失败：" + ex.Message;
                    continue;
                }
                if (cr == null || !cr.Found) continue;      // 定位失败（外部簿打不开等）：跳过
                if (cr.AutoOpened) r.AutoOpened++;

                // L1 值层：非错误节点 → 剪枝
                bool formulaHasLiteral = FormulaRefParser.ContainsErrorLiteral(cr.Formula);
                if (!cr.HasErrorValue && !formulaHasLiteral) continue;

                // L0 文本层：公式含错误字面量 → 错误源（不再回溯）
                if (formulaHasLiteral)
                {
                    AddSource(r, cur, cr, ErrorSourceReason.LiteralInFormula, options);
                    if (r.Sources.Count >= options.MaxSources)
                    {
                        r.TruncatedByLimit = true;
                        r.LimitNote = "达到错误源上限 " + options.MaxSources;
                        break;
                    }
                    continue;
                }

                // 到这里：HasErrorValue=true 且公式无错误字面量
                // 无公式（理论上错误值必有公式，兜底按自算错误源处理）
                if (string.IsNullOrEmpty(cr.Formula) || !cr.Formula.StartsWith("="))
                {
                    AddSource(r, cur, cr, ErrorSourceReason.SelfComputed, options);
                    continue;
                }

                // L2 引用层：展开上游，筛查错误格入队
                if (cur.Depth >= options.MaxDepth)
                {
                    // 到深度上限仍有错误但未定位源头：按自算错误兜底收录
                    r.TruncatedByLimit = true;
                    if (r.LimitNote.Length == 0) r.LimitNote = "达到深度上限 " + options.MaxDepth;
                    AddSource(r, cur, cr, ErrorSourceReason.SelfComputed, options);
                    continue;
                }

                List<PrecedentItem> refs = FormulaRefParser.ParseAll(cr.Formula);
                bool anyErrorUpstream = false;
                for (int i = 0; i < refs.Count; i++)
                {
                    PrecedentItem p = refs[i];
                    // 补全同表引用的簿/表（ParseAll 不填）
                    string pWbPath = p.WorkbookPath, pWbName = p.WorkbookName, pSheet = p.SheetName;
                    if (p.RefType == RefType.SameSheet)
                    { pWbPath = cur.WbPath; pWbName = cur.WbName; pSheet = cur.Sheet; }
                    else if (p.RefType == RefType.CrossSheet)
                    { pWbPath = cur.WbPath; pWbName = cur.WbName; }

                    // 区域引用展开首格（M1 简化：只查区域首格，与原追踪引用一致）
                    CellReadResult up;
                    try { up = reader(pWbPath, pWbName, pSheet, p.CellAddressA1); }
                    catch { continue; }
                    if (up == null || !up.Found) continue;
                    if (up.AutoOpened) r.AutoOpened++;

                    bool upErr = up.HasErrorValue || FormulaRefParser.ContainsErrorLiteral(up.Formula);
                    if (!upErr) continue;                   // 上游正常 → 不入队
                    anyErrorUpstream = true;

                    QNode next = new QNode();
                    next.WbPath = pWbPath ?? ""; next.WbName = pWbName ?? "";
                    next.Sheet = pSheet ?? ""; next.Addr = p.CellAddressA1;
                    next.Depth = cur.Depth + 1;
                    next.PathText = cur.PathText + " ← " + p.CellAddressA1;
                    next.IsCrossWorkbook = p.RefType == RefType.CrossWorkbook;
                    string key = NodeKey(next);
                    if (visited.Add(key)) queue.Enqueue(next);   // 入队时写入 visited 防环
                }

                // 有错误上游 → 传播节点，继续回溯；全部上游正常 → 自算错误源
                if (!anyErrorUpstream)
                    AddSource(r, cur, cr, ErrorSourceReason.SelfComputed, options);

                if (r.Sources.Count >= options.MaxSources)
                {
                    r.TruncatedByLimit = true;
                    r.LimitNote = "达到错误源上限 " + options.MaxSources;
                    break;
                }
            }

            r.Success = true;
            return r;
        }

        static void AddSource(ErrorTraceResult r, QNode cur, CellReadResult cr,
            ErrorSourceReason reason, ErrorTraceOptions options)
        {
            if (r.Sources.Count >= options.MaxSources) return;
            ErrorSourceItem item = new ErrorSourceItem();
            item.WorkbookPath = cur.WbPath;
            item.WorkbookName = cur.WbName;
            item.SheetName = cur.Sheet;
            item.CellAddressA1 = cur.Addr;
            item.Formula = cr.Formula ?? "";
            item.ErrorText = cr.ErrorText ?? "";
            item.Reason = reason;
            item.Depth = cur.Depth;
            item.PathText = cur.PathText;
            item.IsCrossWorkbook = cur.IsCrossWorkbook;
            r.Sources.Add(item);
        }

        static string NodeKey(QNode n)
        {
            // visited 键：簿路径（无路径用簿名）+表+地址，大写归一
            string wb = n.WbPath.Length > 0 ? n.WbPath : n.WbName;
            return (wb + "|" + n.Sheet + "|" + n.Addr).ToUpperInvariant();
        }

        // ---- TraceFrom：碰 COM 的入口 ----

        /// <summary>
        /// 从活动上下文启动回溯。rootWb*/rootSheet/rootAddrs 由调用方（TrackerForm）给出，
        /// reader 用 ComCellReader（读 Value2/Value/Text，必要时自动开簿）。
        /// </summary>
        public static ErrorTraceResult TraceFrom(object application,
            string rootWbPath, string rootWbName, string rootSheet, IList<string> rootAddrs,
            ErrorTraceOptions options, Func<bool> shouldCancel)
        {
            if (options == null) options = new ErrorTraceOptions();
            ComCellReader com = new ComCellReader(application, options);
            ErrorTraceResult r = TraceCore(rootWbPath, rootWbName, rootSheet, rootAddrs,
                com.Read, options, shouldCancel);
            r.AutoOpened = com.AutoOpened;
            if (com.HitAutoOpenLimit)
            {
                r.TruncatedByLimit = true;
                if (r.LimitNote.Length == 0) r.LimitNote = "达到自动开簿上限 " + options.MaxAutoOpen;
            }
            return r;
        }

        /// <summary>
        /// 读单格（供 TrackerForm 行级错误判定复用）。
        /// autoOpen=false 时外部工作簿未打开则 Found=false，不做任何打开动作。
        /// </summary>
        public static CellReadResult ReadCell(object application,
            string wbPath, string wbName, string sheetName, string addr, bool autoOpen)
        {
            ErrorTraceOptions opt = new ErrorTraceOptions();
            opt.AutoOpenWorkbook = autoOpen;
            ComCellReader com = new ComCellReader(application, opt);
            return com.Read(wbPath ?? "", wbName ?? "", sheetName ?? "", addr ?? "");
        }

        // ---- ComCellReader：COM 读格（私有，碰 COM 唯一处）----

        sealed class ComCellReader
        {
            readonly dynamic _app;
            readonly ErrorTraceOptions _opt;
            public int AutoOpened;
            public bool HitAutoOpenLimit;

            public ComCellReader(object app, ErrorTraceOptions opt)
            { _app = app; _opt = opt; }

            public CellReadResult Read(string wbPath, string wbName, string sheetName, string addr)
            {
                CellReadResult r = new CellReadResult();
                try
                {
                    dynamic wb = PrecedentTrackerEngine.FindOpenWorkbook(_app, wbName, wbPath);
                    if (wb == null && _opt.AutoOpenWorkbook && !string.IsNullOrEmpty(wbPath))
                    {
                        if (AutoOpened >= _opt.MaxAutoOpen) { HitAutoOpenLimit = true; r.FailNote = "开簿上限"; return r; }
                        if (!File.Exists(wbPath)) { r.FailNote = "文件不存在"; return r; }
                        wb = _app.Workbooks.Open(wbPath);
                        AutoOpened++;
                        r.AutoOpened = true;
                    }
                    if (wb == null) { r.FailNote = "工作簿未打开"; return r; }

                    dynamic ws = PrecedentTrackerEngine.FindSheet(wb, sheetName);
                    if (ws == null) { r.FailNote = "工作表不存在"; return r; }

                    dynamic cell = ws.Range[addr];
                    r.Found = true;

                    try { r.Formula = (string)cell.Formula; } catch { r.Formula = ""; }

                    // L1 值层：Value2 → Value → Text 三级读错误
                    object v = null;
                    try { v = cell.Value2; } catch { }
                    string err = TryReadError(v);
                    if (err == null)
                    {
                        try { v = cell.Value; } catch { v = null; }
                        err = TryReadError(v);
                    }
                    if (err == null)
                    {
                        try { v = cell.Text; } catch { v = null; }
                        err = TryReadError(v);
                    }
                    r.HasErrorValue = err != null;
                    r.ErrorText = err ?? "";
                    return r;
                }
                catch (Exception ex)
                {
                    r.Found = false;
                    r.FailNote = ex.Message;
                    return r;
                }
            }
        }
    }
}
