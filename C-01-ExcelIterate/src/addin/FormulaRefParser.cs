using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    // =====================================================================
    // 公式引用解析核心（纯字符串解析，不碰 COM）
    //
    // 分层设计：
    //   FormulaRefParser 只做公式字符串 → PrecedentItem 列表的纯解析。
    //   COM 交互（查已开簿、定义名解析）由 PrecedentTrackerEngine 负责。
    //   这样 TrackerProbe 可以直接喂字符串离线自测，不需要真实 WPS 表格。
    //
    // 支持 4 种引用形态：
    //   形态1 同表：O8 / A1:B2 / $A$1
    //   形态2 跨表无引号：单体数据汇总!F377 / Notes!G1118
    //   形态3 跨表带引号：'word附注链接（勿看、勿删）'!F712
    //   形态4 跨簿带路径：'D:\...\[工作簿1]Sheet1'!$J$8
    // =====================================================================

    public enum TrackerStatus { NoWorkbook, NoActiveCell, NoFormula, Completed, Error }
    public enum RefType { SameSheet, CrossSheet, CrossWorkbook }

    /// <summary>一条被引用单元格的解析结果（追踪引用/从属共用）。</summary>
    public class PrecedentItem
    {
        public int Index;
        public string WorkbookPath = "";      // 跨簿完整路径；同表/跨表由引擎填当前簿 FullName
        public string WorkbookName = "";      // 工作簿文件名
        public string SheetName = "";         // 工作表名（已剥引号）；SameSheet 留空由引擎填
        public string CellAddressA1 = "";     // A1 地址（区域取首格，已去 $）
        public string RangeEndA1 = "";        // 区域末格（已去 $）；单格为空
        public string FormulaFragment = "";   // 该引用在公式里的原文 token
        public RefType RefType;
        public bool IsColAbsolute;            // $ 在列前
        public bool IsRowAbsolute;            // $ 在行前
        public bool IsRange;                  // 是否区域 A1:B2
        public bool IsDefinedName;            // 是否定义名（M2 暂不解析，留 false）
        public string DefinedName = "";
        public string RawQuoted = "";         // 跨簿/跨表引号内原文（调试用）
    }

    /// <summary>追踪引用引擎的输出（含当前单元格信息 + 引用项列表）。</summary>
    public class PrecedentResult
    {
        public TrackerStatus Status;
        public string ErrorMessage;
        public string WorkbookName = "";      // 当前工作簿名
        public string WorkbookPath = "";      // 当前工作簿完整路径（FullName），供「所选单元格」行显示/复制
        public string SheetName = "";         // 当前工作表名
        public string CellAddress = "";       // 当前单元格地址
        public string Formula = "";           // 当前单元格公式原文
        public List<PrecedentItem> Items = new List<PrecedentItem>();
    }

    public static class FormulaRefParser
    {
        // ---- 正则 ----

        // A1 单元格地址：$A$1 / A1 / $A1 / A$1，列 1-3 字母，行数字
        const string CellPattern = @"\$?[A-Za-z]{1,3}\$?\d+";
        // 区域：cell:cell
        const string RangePattern = CellPattern + @"(?::" + CellPattern + @")?";

        // 主匹配（覆盖 4 种形态，先具体后一般）：
        //   A) 引号形态（形态3/4）：'...'!cell[:cell]
        //   B) 无引号跨表/簿（形态2/2b）：[工作簿]工作表!cell 或 工作表!cell
        //   C) 同表（形态1）：cell[:cell]
        static readonly Regex RxRef = new Regex(
            @"^(?:"
            + @"'(?<quoted>[^']*)'!(?<range>" + RangePattern + @")"                    // A
            + @"|"
            + @"(?:(?<path>[^\[\]!]+)?\[(?<wb>[^\]]+)\])?(?<sheet>[^\[\]!]+)!(?<range>" + RangePattern + @")"  // B
            + @"|"
            + @"(?<range>" + RangePattern + @")"                                        // C
            + @")$", RegexOptions.Compiled);

        // 拆分 range 的绝对/相对与区域
        static readonly Regex RxCell = new Regex(
            @"^(?<col1>\$?[A-Za-z]{1,3})(?<row1>\$?\d+)(?::(?<col2>\$?[A-Za-z]{1,3})(?<row2>\$?\d+))?$",
            RegexOptions.Compiled);

        // ---- Tokenize（公式拆 token）----

        /// <summary>
        /// 把公式拆成 operand token 列表。
        /// 规则：去前导 = ；按运算符 +-*/^&amp;=&lt;&gt;(), 拆分；
        /// 单引号串整体保留（跨簿路径可含运算符字符）；双引号字符串整体保留但标记跳过；
        /// 冒号 : 和感叹号 ! 不当分隔符（A1:B2 和 Sheet1!A1 须整体保留）。
        /// </summary>
        internal static List<string> Tokenize(string formula)
        {
            List<string> tokens = new List<string>();
            if (string.IsNullOrEmpty(formula)) return tokens;
            string s = formula.Trim();
            if (s.StartsWith("=")) s = s.Substring(1);

            StringBuilder cur = new StringBuilder();
            bool inSingleQuote = false;
            bool inDoubleQuote = false;
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                if (ch == '\'' && !inDoubleQuote)
                {
                    inSingleQuote = !inSingleQuote;
                    cur.Append(ch);
                    continue;
                }
                if (ch == '"' && !inSingleQuote)
                {
                    inDoubleQuote = !inDoubleQuote;
                    cur.Append(ch);
                    continue;
                }
                if (inSingleQuote || inDoubleQuote)
                {
                    cur.Append(ch);
                    continue;
                }
                // 运算符是分隔符（不含 : ! %）
                if ("+-*/^&=<>(),".IndexOf(ch) >= 0)
                {
                    if (cur.Length > 0) { tokens.Add(cur.ToString()); cur.Clear(); }
                    // 运算符本身不入 token（不是引用）
                }
                else
                {
                    cur.Append(ch);
                }
            }
            if (cur.Length > 0) tokens.Add(cur.ToString());
            return tokens;
        }

        // ---- ParseToken（单 token → PrecedentItem）----

        /// <summary>把单个 token 解析为引用项；非引用返回 null。</summary>
        public static PrecedentItem ParseToken(string token)
        {
            if (string.IsNullOrEmpty(token)) return null;
            // 跳过双引号字符串（不是引用）
            if (token.StartsWith("\"")) return null;

            Match m = RxRef.Match(token);
            if (!m.Success) return null;

            PrecedentItem item = new PrecedentItem();
            item.FormulaFragment = token;

            // --- 形态 A：引号形态 ---
            if (m.Groups["quoted"].Success)
            {
                string quoted = m.Groups["quoted"].Value;
                item.RawQuoted = quoted;
                int lb = quoted.IndexOf('[');
                int rb = lb >= 0 ? quoted.IndexOf(']', lb + 1) : -1;
                if (lb >= 0 && rb > lb)
                {
                    // 形态4：跨簿带路径 '...路径...[工作簿]工作表'!cell
                    item.WorkbookPath = quoted.Substring(0, lb);
                    item.WorkbookName = quoted.Substring(lb + 1, rb - lb - 1);  // 可能含扩展名也可能不含
                    item.SheetName = quoted.Substring(rb + 1);
                    item.RefType = RefType.CrossWorkbook;
                }
                else
                {
                    // 形态3：跨表带引号 '工作表'!cell
                    item.SheetName = quoted;
                    item.RefType = RefType.CrossSheet;
                }
            }
            // --- 形态 B：无引号跨表/簿 ---
            else if (m.Groups["sheet"].Success && m.Groups["sheet"].Value.Length > 0)
            {
                if (m.Groups["wb"].Success && m.Groups["wb"].Value.Length > 0)
                {
                    // 形态2b：[工作簿]工作表!cell（罕见，Excel 通常自动加引号）
                    item.WorkbookPath = m.Groups["path"].Success ? m.Groups["path"].Value : "";
                    item.WorkbookName = m.Groups["wb"].Value;
                    item.SheetName = m.Groups["sheet"].Value;
                    item.RefType = RefType.CrossWorkbook;
                }
                else
                {
                    // 形态2：跨表无引号 工作表!cell
                    item.SheetName = m.Groups["sheet"].Value;
                    item.RefType = RefType.CrossSheet;
                }
            }
            // --- 形态 C：同表 ---
            else
            {
                item.RefType = RefType.SameSheet;
                // SheetName 留空，由 PrecedentTrackerEngine 填当前活动表
            }

            // 解析 range 部分（绝对/相对 + 区域）
            ParseRange(m.Groups["range"].Value, item);
            return item;
        }

        /// <summary>拆分 range 字符串的绝对/相对标识与区域范围。</summary>
        private static void ParseRange(string rangeStr, PrecedentItem item)
        {
            Match m = RxCell.Match(rangeStr);
            if (!m.Success) return;
            string col1 = m.Groups["col1"].Value;
            string row1 = m.Groups["row1"].Value;
            item.IsColAbsolute = col1.StartsWith("$");
            item.IsRowAbsolute = row1.StartsWith("$");
            item.CellAddressA1 = col1.TrimStart('$') + row1.TrimStart('$');
            if (m.Groups["col2"].Success)
            {
                item.IsRange = true;
                item.RangeEndA1 = m.Groups["col2"].Value.TrimStart('$')
                                + m.Groups["row2"].Value.TrimStart('$');
            }
        }

        // ---- ParseAll（公式 → 全部引用项）----

        /// <summary>解析公式中的全部引用项（纯字符串解析，不碰 COM）。</summary>
        public static List<PrecedentItem> ParseAll(string formula)
        {
            List<PrecedentItem> items = new List<PrecedentItem>();
            List<string> tokens = Tokenize(formula);
            int idx = 0;
            for (int i = 0; i < tokens.Count; i++)
            {
                PrecedentItem item = ParseToken(tokens[i]);
                if (item == null) continue;
                item.Index = ++idx;
                items.Add(item);
            }
            return items;
        }

        // ---- 错误字面量检测（「显示错误来源」的文本层判据，纯字符串，不碰 COM）----

        // Excel 错误字面量（公式文本里残留的形态；忽略大小写）。
        static readonly string[] ErrorLiterals = { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A" };

        /// <summary>
        /// 剔除公式中的双引号字符串字面量（含引号本身），只保留引号外的部分。
        /// 用途：判错误字面量时不把 ="提示：#REF!" 这类显示文本误判为错误引用。
        /// Excel 字符串内用 "" 表示一个引号，剔除时按成对跳过。
        /// </summary>
        public static string StripStringLiterals(string formula)
        {
            if (string.IsNullOrEmpty(formula)) return formula ?? "";
            StringBuilder sb = new StringBuilder(formula.Length);
            bool inDoubleQuote = false;
            for (int i = 0; i < formula.Length; i++)
            {
                char ch = formula[i];
                if (ch == '"')
                {
                    // 字符串内的 "" 是转义的一个引号：跳过这一对，不切换状态
                    if (inDoubleQuote && i + 1 < formula.Length && formula[i + 1] == '"') { i++; continue; }
                    inDoubleQuote = !inDoubleQuote;
                    continue;                    // 引号本身也不保留
                }
                if (!inDoubleQuote) sb.Append(ch);
            }
            return sb.ToString();
        }

        /// <summary>公式文本（剔除双引号字符串后）是否含错误字面量（#REF! 等）。</summary>
        public static bool ContainsErrorLiteral(string formula)
        { return FirstErrorLiteral(formula) != null; }

        /// <summary>返回公式文本（剔除双引号字符串后）命中的第一个错误字面量；无则 null。</summary>
        public static string FirstErrorLiteral(string formula)
        {
            string s = StripStringLiterals(formula);
            if (s.Length == 0) return null;
            for (int i = 0; i < ErrorLiterals.Length; i++)
                if (s.IndexOf(ErrorLiterals[i], StringComparison.OrdinalIgnoreCase) >= 0)
                    return ErrorLiterals[i];
            return null;
        }
    }
}
