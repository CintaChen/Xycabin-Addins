using System;
using System.Collections.Generic;
using System.Globalization;

namespace Excel_Iterate
{
    // =====================================================================
    // 合并抵消引擎（「往来抵消 / 收入成本抵消」核心计算，纯逻辑）
    //
    // 会计模型（与设计备忘一致，金额单位默认为元）：
    //   往来抵消（按单位聚合多行后逐单位出分录）：
    //     对冲额 = min(应收, 应付)
    //     分录① 借 应付账款 / 贷 应收账款（对冲额）
    //     当 应收>应付（对冲后仍有剩余应收）且数据含坏账准备列时：
    //     分录② 借 坏账准备 / 贷 信用减值损失
    //            金额 = 坏账准备 × 对冲额 ÷ 应收（按对冲比例冲回已提准备）
    //   收入成本抵消（逐行）：
    //     未实现利润 U = (R − C) × 未对外销售比例
    //     分录 借 营业收入 R；贷 营业成本 (R − U)；贷 存货 U
    //     （借贷平衡：R = (R−U) + U；比例=0 时全额反转收入成本，
    //       比例=100% 时贷营业成本=C、贷存货=R−C）
    //
    // 分层设计：
    //   EliminationEngine 纯核：只吃 EliminationRow（COM 层读好的文本），
    //     吐 EliminationResult；任意线程可调，零 COM 零 WinForms，
    //     探针（EliminationProbe）可注入假数据离线自测。
    //   EliminationHost / EliminationForm：碰 COM 与 UI 的层。
    // =====================================================================

    /// <summary>抵消功能种类。</summary>
    public enum EliminationKind
    {
        /// <summary>往来抵消（内部债权债务对冲 + 坏账准备按比例冲回）。</summary>
        Receivable,
        /// <summary>收入成本抵消（内部交易收入成本与未实现利润）。</summary>
        RevenueCost
    }

    /// <summary>数据列在抵消业务中承担的角色（列映射网格用）。</summary>
    public enum EliminationRole
    {
        /// <summary>往来单位 / 客户名称（聚合键）。</summary>
        Unit,
        /// <summary>应收（借方余额）。</summary>
        Receivable,
        /// <summary>应付（贷方余额）。</summary>
        Payable,
        /// <summary>坏账准备（已提准备，可选列）。</summary>
        BadDebt,
        /// <summary>营业收入 R。</summary>
        Revenue,
        /// <summary>营业成本 C。</summary>
        Cost,
        /// <summary>未对外销售比例（0~1，支持 "30%" 写法）。</summary>
        UnsoldRatio,
        /// <summary>核算主体（矩阵式底稿的分组键之一）。</summary>
        Entity,
        /// <summary>客商 / 二级明细（矩阵式底稿的分组键之一）。</summary>
        Counterparty,
        /// <summary>年份 / 会计期间（矩阵式底稿的年度列来源）。</summary>
        Year,
        /// <summary>参与抵消的条件列（如「是否为合并内关联方」，值需等于设定的期望值）。</summary>
        Condition,
        /// <summary>未实现内部交易损益（可选列；映射且本组有值时优先用它，否则按 (收入−成本)×比例 现算）。</summary>
        UnrealizedProfit,
        /// <summary>会计科目（往来抵消：用于区分资产科目与负债科目）。</summary>
        Account,
        /// <summary>审定数（往来抵消的取数金额列）。</summary>
        Amount,
        /// <summary>截止日（往来抵消的年度来源：取前 4 位作年度，如 20260630 → 2026）。</summary>
        Deadline,
        /// <summary>不参与计算的列。</summary>
        Ignore
    }

    /// <summary>一列数据的映射信息（表头行读出的列 + 用户指定角色）。</summary>
    public class EliminationColumn
    {
        public int Number = 1;                 // 列号（1 起）
        public string Letter = "A";            // 列字母
        public string Header = "";             // 表头文本
        public EliminationRole Role = EliminationRole.Ignore;
    }

    /// <summary>一行源数据（COM 层按角色取值后的文本行）。</summary>
    public class EliminationRow
    {
        public int SourceRow = 0;              // 源工作表行号（1 起，供提示与定位）
        public Dictionary<EliminationRole, string> Values = new Dictionary<EliminationRole, string>();
    }

    /// <summary>分录中的一行（借或贷）。金额为 0 的行不入结果。</summary>
    public class EliminationLine
    {
        public bool IsDebit;                   // true=借方行，false=贷方行
        public string Account = "";            // 科目名
        public decimal Amount;                 // 金额（两位小数）
        public string Note = "";               // 计算说明
    }

    /// <summary>一条抵消分录（可含多行借贷，组内借贷必须相等）。</summary>
    public class EliminationEntry
    {
        public int Index;                      // 分录序号（1 起）
        public string Group = "";              // 聚合键（往来=单位名；收入成本=来源行号）
        public string Summary = "";            // 摘要
        public List<EliminationLine> Lines = new List<EliminationLine>();
        public decimal DebitTotal              // 借方合计
        { get { decimal t = 0; foreach (EliminationLine l in Lines) if (l.IsDebit) t += l.Amount; return t; } }
        public decimal CreditTotal             // 贷方合计
        { get { decimal t = 0; foreach (EliminationLine l in Lines) if (!l.IsDebit) t += l.Amount; return t; } }
    }

    /// <summary>抵消计算结果（预检通过才有 Entries；Errors 非空即阻断执行）。</summary>
    public class EliminationResult
    {
        public List<EliminationEntry> Entries = new List<EliminationEntry>();
        public List<string> Warnings = new List<string>();
        public List<string> Errors = new List<string>();
        public int SkippedRows;                // 因缺列/零余额/无法解析而跳过的行数
        public decimal TotalDebit
        { get { decimal t = 0; foreach (EliminationEntry e in Entries) t += e.DebitTotal; return t; } }
        public decimal TotalCredit
        { get { decimal t = 0; foreach (EliminationEntry e in Entries) t += e.CreditTotal; return t; } }
        public bool Balanced                   // 每条分录自身平衡且总借=总贷
        {
            get
            {
                foreach (EliminationEntry e in Entries) if (e.DebitTotal != e.CreditTotal) return false;
                return TotalDebit == TotalCredit;
            }
        }
    }

    /// <summary>合并抵消引擎入口（纯逻辑；UI 线程/探针均可调用）。</summary>
    public static partial class EliminationEngine
    {
        // ---- 科目名（固定，写入门分录；改科目只改这里）----
        public const string AccPayable = "应付账款";
        public const string AccReceivable = "应收账款";
        public const string AccBadDebt = "坏账准备";
        public const string AccCreditLoss = "信用减值损失";
        public const string AccRevenue = "营业收入";
        public const string AccCost = "营业成本";
        public const string AccInventory = "存货";

        // ---- 金额/比例解析 ----

        /// <summary>解析金额文本：容忍千分位逗号、空格、¥/￥/元 前缀后缀、括号负数。</summary>
        public static bool TryParseAmount(string text, out decimal value)
        {
            value = 0m;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string s = text.Trim();
            bool negative = false;
            if (s.Length >= 2 && s[0] == '(' && s[s.Length - 1] == ')')
            {
                negative = true;
                s = s.Substring(1, s.Length - 2).Trim();
            }
            foreach (char c in new[] { ',', ' ', '\u00a0', '\uffe5', '\u00a5', '\u5143', '$' })
                s = s.Replace(c.ToString(), "");
            if (s.StartsWith("-", StringComparison.Ordinal)) { negative = !negative; s = s.Substring(1); }
            else if (s.StartsWith("+", StringComparison.Ordinal)) s = s.Substring(1);
            if (s.Length == 0) return false;
            if (!decimal.TryParse(s, NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign,
                CultureInfo.InvariantCulture, out value))
                return false;
            if (negative) value = -value;
            return true;
        }

        /// <summary>
        /// 解析未对外销售比例："30%" → 0.3；"0.25" → 0.25（裸数字按小数计）。
        /// 越界（&lt;0 或 &gt;1）拒绝——裸数字写 30 会被当成 30 拒绝，避免歧义。
        /// </summary>
        public static bool TryParseRatio(string text, out decimal ratio)
        {
            ratio = 0m;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string s = text.Trim().Replace(" ", "").Replace("\u00a0", "");
            bool percent = s.EndsWith("%", StringComparison.Ordinal);
            if (percent) s = s.Substring(0, s.Length - 1);
            if (s.Length == 0) return false;
            if (!decimal.TryParse(s, NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign,
                CultureInfo.InvariantCulture, out ratio))
                return false;
            if (percent) ratio = ratio / 100m;
            if (ratio < 0m || ratio > 1m) { ratio = 0m; return false; }
            return true;
        }

        // ---- 角色猜测与校验 ----

        /// <summary>
        /// 按表头文本猜角色（用户可在列映射网格覆盖）。
        /// 收入成本抵消走矩阵式底稿，角色集合与往来抵消不同，故按 kind 分支：
        /// 矩阵式必须先认「核算主体 / 客商 / 年份 / 条件列 / 损益列」，
        /// 否则「营业成本」这类表头会被旧规则误判成单位列。
        /// </summary>
        public static EliminationRole GuessRole(string header, EliminationKind kind)
        {
            if (string.IsNullOrWhiteSpace(header)) return EliminationRole.Ignore;
            string h = header.Trim();

            if (kind == EliminationKind.RevenueCost)
            {
                if (h.Contains("未实现") && h.Contains("损益")) return EliminationRole.UnrealizedProfit;
                if (h.Contains("内部交易损益")) return EliminationRole.UnrealizedProfit;
                if (h.Contains("核算主体") || h.Contains("报表主体")
                    || (h.Contains("主体") && !h.Contains("客商"))) return EliminationRole.Entity;
                if (h.Contains("客商") || h.Contains("客户名称") || h.Contains("对方单位")
                    || h.Contains("交易对方") || h.Contains("二级明细")) return EliminationRole.Counterparty;
                if (h.Contains("年份") || h.Contains("年度") || h.Contains("会计期间")) return EliminationRole.Year;
                if (h.Contains("合并内关联方")) return EliminationRole.Condition;
                if (h.Contains("未对外销售") || h.Contains("销售比例")) return EliminationRole.UnsoldRatio;
                if (h.Contains("单体收入")) return EliminationRole.Revenue;
                if (h.Contains("单体成本")) return EliminationRole.Cost;
                return EliminationRole.Ignore;
            }

            // 往来抵消（Receivable）：先认审定数/截止日/科目这类强特征列，
            // 否则「应收」「应付」会被科目列里的科目名误判成余额列。
            if (h.Contains("审定数") || h.Contains("审定金额")) return EliminationRole.Amount;
            if (h.Contains("截止日") || h.Contains("基准日")) return EliminationRole.Deadline;
            if (h.Contains("科目")) return EliminationRole.Account;
            if (h.Contains("核算主体") || h.Contains("报表主体")) return EliminationRole.Entity;
            if (h.Contains("对方单位") || h.Contains("客商")) return EliminationRole.Counterparty;
            if (h.Contains("合并内关联方")) return EliminationRole.Condition;

            if (h.Contains("坏账")) return EliminationRole.BadDebt;
            if (h.Contains("应收")) return EliminationRole.Receivable;
            if (h.Contains("应付")) return EliminationRole.Payable;
            if (h.Contains("未对外销售") || h.Contains("未销售比例")
                || (h.Contains("销售") && h.Contains("比例"))) return EliminationRole.UnsoldRatio;
            if (h.Contains("收入")) return EliminationRole.Revenue;
            if (h.Contains("成本")) return EliminationRole.Cost;
            if (h.Contains("单位") || h.Contains("客户") || h.Contains("往来")
                || h.Contains("公司") || h.Contains("交易") || h.Contains("内容")) return EliminationRole.Unit;
            return EliminationRole.Ignore;
        }

        /// <summary>
        /// 该种类下必须映射的列角色。
        /// 往来抵消：单位 + 应收 + 应付（坏账准备可选）。
        /// 收入成本抵消（矩阵式）：分组三维（核算主体、客商、年份）+ 收入 + 成本；
        /// 未实现损益的两个来源（损益列 / 未对外销售比例列）至少要有其一，由 ValidateMatrixMapping 校验。
        /// </summary>
        public static IList<EliminationRole> RequiredRoles(EliminationKind kind)
        {
            if (kind == EliminationKind.Receivable)
                return new[] { EliminationRole.Entity, EliminationRole.Counterparty,
                               EliminationRole.Account, EliminationRole.Amount,
                               EliminationRole.Deadline };
            return new[] { EliminationRole.Entity, EliminationRole.Counterparty,
                           EliminationRole.Year, EliminationRole.Revenue, EliminationRole.Cost };
        }

        /// <summary>该映射是否包含指定角色。</summary>
        public static bool HasRole(IList<EliminationColumn> columns, EliminationRole role)
        {
            if (columns == null) return false;
            foreach (EliminationColumn column in columns)
                if (column.Role == role) return true;
            return false;
        }

        /// <summary>
        /// 矩阵式底稿的补充校验：未实现内部交易损益必须有来源——
        /// 映射「未实现内部交易损益」列，或映射「未对外销售比例」列由程序现算，二者至少其一。
        /// </summary>
        public static string ValidateMatrixMapping(EliminationKind kind, IList<EliminationColumn> columns)
        {
            // 往来抵消不涉及收入成本口径，无需校验损益来源
            if (kind == EliminationKind.Receivable) return null;
            if (HasRole(columns, EliminationRole.UnrealizedProfit)) return null;
            if (HasRole(columns, EliminationRole.UnsoldRatio)) return null;
            return "未实现内部交易损益无法计算：请映射「未实现内部交易损益」列，"
                + "或映射「未对外销售比例」列（二者至少其一）。";
        }

        /// <summary>校验列映射：缺必填角色或同一角色指定两列时返回错误文本，否则 null。</summary>
        public static string ValidateMapping(EliminationKind kind, IList<EliminationColumn> columns)
        {
            Dictionary<EliminationRole, EliminationColumn> seen = new Dictionary<EliminationRole, EliminationColumn>();
            foreach (EliminationColumn c in columns)
            {
                if (c.Role == EliminationRole.Ignore) continue;
                EliminationColumn prior;
                if (seen.TryGetValue(c.Role, out prior))
                    return "列「" + prior.Letter + " " + prior.Header + "」与「" + c.Letter + " " + c.Header
                        + "」都指定为「" + RoleText(c.Role) + "」，请只保留一列。";
                seen[c.Role] = c;
            }
            foreach (EliminationRole role in RequiredRoles(kind))
                if (!seen.ContainsKey(role))
                    return "缺少必填列：「" + RoleText(role) + "」。请在列映射网格中指定。";
            return null;
        }

        /// <summary>角色中文名（列映射下拉与错误提示用）。</summary>
        public static string RoleText(EliminationRole role)
        {
            switch (role)
            {
                case EliminationRole.Unit: return "单位/客户";
                case EliminationRole.Receivable: return "应收（借方余额）";
                case EliminationRole.Payable: return "应付（贷方余额）";
                case EliminationRole.BadDebt: return "坏账准备";
                case EliminationRole.Revenue: return "营业收入";
                case EliminationRole.Cost: return "营业成本";
                case EliminationRole.UnsoldRatio: return "未对外销售比例";
                case EliminationRole.Entity: return "核算主体";
                case EliminationRole.Counterparty: return "明细（客商）";
                case EliminationRole.Year: return "年份/期间";
                case EliminationRole.Condition: return "参与抵消条件列";
                case EliminationRole.UnrealizedProfit: return "未实现内部交易损益";
                case EliminationRole.Account: return "科目";
                case EliminationRole.Amount: return "审定数";
                case EliminationRole.Deadline: return "截止日/期间";
                default: return "忽略";
            }
        }

        // ---- 分录构建 ----





        // ---- 内部工具 ----

        private sealed class UnitSum
        {
            public decimal Receivable, Payable, BadDebt;
            public int Rows;
        }

        private static string Text(EliminationRow row, EliminationRole role)
        {
            string value;
            return row.Values.TryGetValue(role, out value) ? (value ?? "").Trim() : "";
        }

        private static void Renumber(EliminationResult result)
        {
            for (int i = 0; i < result.Entries.Count; i++) result.Entries[i].Index = i + 1;
        }

        private static decimal Money2(decimal value)
        { return Math.Round(value, 2, MidpointRounding.AwayFromZero); }

        private static string Money(decimal value)
        { return Money2(value).ToString("0.00", CultureInfo.InvariantCulture); }

        private static string RatioText(decimal ratio)
        { return (ratio * 100m).ToString("0.##", CultureInfo.InvariantCulture) + "%"; }
    }
}
