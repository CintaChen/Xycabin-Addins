using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    // =====================================================================
    // 合并抵消（收入成本）—— 矩阵式审计底稿纯核
    //
    // 输出版式（与「收入抵消结果」样张一致）：
    //   A 核算主体 | B 明细【客商】 | C 备注 | D 调整事项 | E 会计科目
    //   之后每个年度占三列：[空分隔列][借方][贷方]，年度按源数据实际出现值降序
    //   第 1 行（可选）SUBTOTAL 合计；第 2/3 行双层表头
    //
    // 每个（核算主体 → 客商）交易对占固定的一组行，行内容由「行模板」决定：
    //   行1 借 营业收入   R = Σ收入列
    //   行2 贷 营业成本   C = Σ成本列
    //   行3 贷 营业成本   (R − C) − U    （成本还原）
    //   行4 贷 存货       U              （未实现内部交易损益）
    //
    // 恒等式：R ≡ C + [(R−C)−U] + U —— 无条件成立。
    // U 为负（收入 < 成本）时写在贷方列，即实质借方；科目方向固定、金额带符号，
    // 因此不存在做不平的可能，也不必像旧分录模型那样靠 if 守卫取舍存货行。
    //
    // 未实现内部交易损益 U 的两个来源（「损益列可选」）：
    //   ① 映射了「未实现内部交易损益」列且本组该列有值 → Σ该列，写成 SUMIFS
    //   ② 否则 → 逐行 Σ(收入−成本)×未对外销售比例 现算，写成 SUMPRODUCT
    //   注意：现算必须逐行加权，不能先汇总再乘比例（比例可能逐行不同）。
    //
    // 分层：本文件零 COM 零 WinForms，任意线程可调，探针可注入假数据离线自测。
    // 行模板可插拔——第二阶段接入往来抵消只需新增 MatrixTemplates 的一个方法，
    // 无需改动本文件任何计算逻辑（组高 = 模板行数，行号锚点全部推导得出）。
    // =====================================================================

    /// <summary>矩阵底稿的输出选项（窗体收集 → 引擎消费）。</summary>
    public sealed class MatrixOptions
    {
        /// <summary>源工作表表名，用于拼公式前缀「内部交易明细!」。</summary>
        public string SourceSheetName = "";
        /// <summary>源工作簿文件名；空表示与写入目标同簿（不写外部引用）。</summary>
        public string SourceBookName = "";
        /// <summary>条件列的期望值（样张为「是」）。</summary>
        public string ConditionValue = "是";
        /// <summary>是否映射了条件列；false 时全部数据行参与。</summary>
        public bool HasConditionColumn;
        /// <summary>第 1 行是否写 SUBTOTAL 合计。</summary>
        public bool TotalRow = true;
        /// <summary>降级开关：写静态数值而非公式（大数据量或公式核对不稳时用）。</summary>
        public bool StaticValues;
        /// <summary>整列引用（$O:$O）还是限定数据区（$O$3:$O$200）。</summary>
        public bool FullColumnRefs = true;
        /// <summary>年度数量上限，防止误选到日期列产生上百个年度。</summary>
        public int MaxYears = 10;
        /// <summary>源数据区首行（表头行 + 1）。</summary>
        public int SourceFirstRow = 3;
        /// <summary>源数据区末行（UsedRange 底行）；仅 LimitComputedRange=true 时用于限定现算区间。</summary>
        public int SourceLastRow = 0;
        /// <summary>
        /// 现算（SUMPRODUCT）分支是否限定到当前数据区。默认 false——从数据首行算到工作表末尾，
        /// 保证源表新增行后自动纳入；设为 true 则退回 `$O$3:$O$末行`（快，但新增行需重新执行）。
        /// </summary>
        public bool LimitComputedRange;
        /// <summary>写入目标起始行（1 起）。</summary>
        public int StartRow = 1;
        /// <summary>写入目标起始列（1 起）。</summary>
        public int StartColumn = 1;
        /// <summary>
        /// 往来抵消「差异表」的实际表名。结果表的金额以跨表引用差异表的「抵消金额」格，
        /// 因此需要真实表名（重名时会被自动编号，故执行阶段必须以建表后的实际名为准）。
        /// </summary>
        public string DiffSheetName = "";
        /// <summary>
        /// 每个交易对（组）之间空几行；0=不空行。这些行留在写入块内并写 null，
        /// 与「覆盖写入」语义一致（PrepareTargetRegion 本就会 ClearContents 整个矩形）。
        /// </summary>
        public int GroupGap;
        /// <summary>一组占几行（= 行模板条数）。</summary>
        public int GroupHeight;
        /// <summary>组间距生效后的每步跨行数（= GroupHeight + GroupGap），行号推导统一用它。</summary>
        internal int GroupStride()
        { int h = GroupHeight > 0 ? GroupHeight : 1; return h + (GroupGap > 0 ? GroupGap : 0); }
    }

    /// <summary>一个金额格：既能出静态值（降级/断言用），也能出活公式（默认写入用）。</summary>
    public sealed class MatrixAmount
    {
        public decimal Value;
        public string Formula;

        /// <summary>写入单元格时真正落下去的东西：默认公式，降级模式为数值。</summary>
        public static object Cell(MatrixAmount amount, bool staticValues)
        {
            if (amount == null) return null;
            if (!staticValues && amount.Formula != null && amount.Formula.Length > 0)
                return amount.Formula;
            return Convert.ToDouble(amount.Value);
        }
    }

    /// <summary>行模板中一项求和成分的来源。</summary>
    public enum MatrixCellKind
    {
        /// <summary>引用源表度量的 SUMIFS / SUMPRODUCT。</summary>
        Measure,
        /// <summary>引用本组另一行已写出的单元格（如 =(G4-H5)-H7）。</summary>
        RowRef
    }

    /// <summary>行模板中的一项（带正负号），多项按顺序拼成表达式。</summary>
    public sealed class MatrixTerm
    {
        public bool Plus;
        public MatrixCellKind Kind = MatrixCellKind.Measure;
        /// <summary>Measure 时为度量键；RowRef 时为被引用模板行的 Key。</summary>
        public string Key = "";
    }

    /// <summary>
    /// 一组交易对中的一行模板。可插拔的核心抽象：
    /// 改科目、改备注、改计算式都只动这里，构建器与公式器具一概不动。
    /// </summary>
    public sealed class MatrixRowTemplate
    {
        public string Key = "";
        /// <summary>备注模板，支持 {Entity} {Customer} 占位。</summary>
        public string NoteTemplate = "";
        public string Adjustment = "";
        public string Account = "";
        /// <summary>true 写年度对的借方列；false 写贷方列。</summary>
        public bool IsDebit;
        public List<MatrixTerm> Terms = new List<MatrixTerm>();
        /// <summary>
        /// 可选的表达式模板，形如 "=({0}-{1})-{2}"，{i} 按 Terms 顺序代入。
        /// 需要明确分组（如「（收入−成本）−未实现损益」）时用；留空则按 Terms 顺序左结合拼接。
        /// 模板必须自带前导 "="，否则不会被识别为公式。
        /// </summary>
        public string ExpressionTemplate = "";
    }



    /// <summary>输出表的一列（固定五列 + 每年度三列）。</summary>
    public sealed class MatrixColumn
    {
        /// <summary>绝对列号（1 起）。</summary>
        public int Number;
        /// <summary>绝对列字母。</summary>
        public string Letter = "";
        /// <summary>块内偏移（0 起），写入计划与合并表头用。</summary>
        public int Index;
        /// <summary>分隔列（年度之间的空列）。</summary>
        public bool IsSpacer;
        /// <summary>年度文本；固定列与分隔列为 null。</summary>
        public string Year;
        /// <summary>金额列的借贷方向。</summary>
        public bool IsDebit;
        /// <summary>第 2 行表头文本。</summary>
        public string Head2;
        /// <summary>第 3 行表头文本（借/贷）；null 表示参与纵向合并。</summary>
        public string Head3;
    }

    /// <summary>输出表的一行。</summary>
    public sealed class MatrixRow
    {
        public int RowNumber;
        public string Entity = "";
        public string Customer = "";
        public string Note = "";
        public string Adjustment = "";
        public string Account = "";
        /// <summary>生成该行的模板 Key（定位与断言用）。</summary>
        public string TemplateKey = "";
        public Dictionary<string, MatrixAmount> Debit = new Dictionary<string, MatrixAmount>();
        public Dictionary<string, MatrixAmount> Credit = new Dictionary<string, MatrixAmount>();
    }

    /// <summary>矩阵底稿计算结果。</summary>
    public sealed class EliminationMatrix
    {
        public List<string> Years = new List<string>();
        public List<MatrixColumn> Columns = new List<MatrixColumn>();
        public List<MatrixRow> Rows = new List<MatrixRow>();
        public List<string> Warnings = new List<string>();
        public List<string> Errors = new List<string>();
        public int GroupCount;
        public int MatchedRows;
        public int SkippedRows;
        public int FilteredRows;
        /// <summary>每个交易对占几行（= 行模板条数），不硬编码为 4。</summary>
        public int TemplateCount;
        /// <summary>交易对之间空几行（由 MatrixOptions.GroupGap 带入）。</summary>
        public int GroupGap;
        /// <summary>组间跨行数 = TemplateCount + GroupGap；行号推导统一用它，避免三处公式不一致。</summary>
        public int GroupStride;

        public int ColumnCount { get { return Columns.Count; } }
        public int RowCount { get { return Rows.Count; } }

        public MatrixColumn DebitColumn(string year)
        {
            foreach (MatrixColumn column in Columns)
                if (!column.IsSpacer && column.IsDebit && column.Year == year) return column;
            return null;
        }

        public MatrixColumn CreditColumn(string year)
        {
            foreach (MatrixColumn column in Columns)
                if (!column.IsSpacer && !column.IsDebit && column.Year == year) return column;
            return null;
        }

        /// <summary>每组每年 Σ借 == Σ贷（容差 0.005，吸收分位舍入）。</summary>
        public bool Balanced
        {
            get
            {
                if (GroupCount == 0 || Rows.Count == 0 || TemplateCount == 0) return true;
                int dataTop = Rows[0].RowNumber;
                int stride = GroupStride > 0 ? GroupStride : TemplateCount;
                for (int g = 0; g < GroupCount; g++)
                {
                    EliminationCancellation.Check();
                    int top = dataTop + g * stride;
                    for (int y = 0; y < Years.Count; y++)
                    {
                        EliminationCancellation.Check();
                        decimal debit = 0m, credit = 0m;
                        for (int r = 0; r < Rows.Count; r++)
                        {
                            EliminationCancellation.Check();
                            MatrixRow row = Rows[r];
                            if (row.RowNumber < top || row.RowNumber >= top + TemplateCount) continue;
                            MatrixAmount amount;
                            if (row.Debit.TryGetValue(Years[y], out amount)) debit += amount.Value;
                            if (row.Credit.TryGetValue(Years[y], out amount)) credit += amount.Value;
                        }
                        if (Math.Abs(debit - credit) > 0.005m) return false;
                    }
                }
                return true;
            }
        }
    }





    /// <summary>源表各角色的列字母（一次扫描缓存，避免反复遍历）。</summary>
    internal sealed class Letters
    {
        internal string Entity, Counterparty, Year, Condition, Revenue, Cost, Ratio, Profit;
        internal string Account, Amount, Deadline;

        internal Letters(IList<EliminationColumn> columns)
        {
            if (columns == null) return;
            for (int i = 0; i < columns.Count; i++)
            {
                EliminationCancellation.Check();
                EliminationColumn column = columns[i];
                switch (column.Role)
                {
                    case EliminationRole.Entity: Entity = column.Letter; break;
                    case EliminationRole.Counterparty: Counterparty = column.Letter; break;
                    case EliminationRole.Year: Year = column.Letter; break;
                    case EliminationRole.Condition: Condition = column.Letter; break;
                    case EliminationRole.Revenue: Revenue = column.Letter; break;
                    case EliminationRole.Cost: Cost = column.Letter; break;
                    case EliminationRole.UnsoldRatio: Ratio = column.Letter; break;
                    case EliminationRole.UnrealizedProfit: Profit = column.Letter; break;
                    case EliminationRole.Account: Account = column.Letter; break;
                    case EliminationRole.Amount: Amount = column.Letter; break;
                    case EliminationRole.Deadline: Deadline = column.Letter; break;
                }
            }
        }
    }



    /// <summary>公式拼装器具：统一处理表名前缀、引用范围与条件转义。</summary>
    internal sealed class FormulaKit
    {
        private readonly MatrixOptions options;

        internal FormulaKit(MatrixOptions options) { this.options = options; }

        /// <summary>源表引用前缀；跨簿时形如 '[文件名.xlsx]表名'!。</summary>
        internal string Prefix()
        {
            string sheet = options.SourceSheetName == null ? "" : options.SourceSheetName;
            if (sheet.Length == 0) return "";
            if (options.SourceBookName != null && options.SourceBookName.Length > 0)
                return "'[" + options.SourceBookName + "]" + sheet + "'!";
            if (NeedsQuote(sheet)) return "'" + sheet + "'!";
            return sheet + "!";
        }

        private static bool NeedsQuote(string sheet)
        {
            for (int i = 0; i < sheet.Length; i++)
            {
                EliminationCancellation.Check();
                char c = sheet[i];
                if (c == ' ' || c == '-' || c == '(' || c == ')' || c == '（' || c == '）') return true;
            }
            return false;
        }

        /// <summary>一列的引用范围（按「引用整列」选项）。</summary>
        internal string Col(string letter)
        {
            if (letter == null || letter.Length == 0) return "";
            return Prefix() + RangeRaw(letter);
        }

        /// <summary>
        /// 现算（SUMPRODUCT）分支专用区间：**从数据首行一直到工作表末尾**（`$O$3:$O$1048576`）。
        /// 这样既避开了第 1/2 行的汇总行与表头文本（整列引用会 #VALUE!，已实测），
        /// 又能随源表向下扩展自动包含新增行——不像「限定到当前数据区」那样写死末行导致新增数据被漏掉。
        /// 代价是扫描行数接近整表；如需性能优先，可把 LimitComputedRange 设为 true 退回当前数据区。
        /// </summary>
        internal string ColData(string letter)
        {
            if (letter == null || letter.Length == 0) return "";
            int last = MergeLimits.MaxRow;
            if (options.LimitComputedRange)
                last = options.SourceLastRow > 0 ? options.SourceLastRow : options.SourceFirstRow;
            return Prefix() + "$" + letter + "$" + options.SourceFirstRow + ":$" + letter + "$" + last;
        }

        private string RangeRaw(string letter)
        {
            if (options.FullColumnRefs) return "$" + letter + ":$" + letter;
            int last = options.SourceLastRow > 0 ? options.SourceLastRow : options.SourceFirstRow;
            return "$" + letter + "$" + options.SourceFirstRow + ":$" + letter + "$" + last;
        }

        internal string Criterion(string text)
        { return "\"" + (text == null ? "" : text.Replace("\"", "\"\"")) + "\""; }

        /// <summary>输出表第 1 列（核算主体）的字母，随起始列平移。</summary>
        private string EntityRef(int rowNumber)
        { return "$" + EliminationMatrixBuilder.Letter(options.StartColumn) + rowNumber; }

        /// <summary>输出表第 2 列（客商）的字母。</summary>
        private string CustomerRef(int rowNumber)
        { return "$" + EliminationMatrixBuilder.Letter(options.StartColumn + 1) + rowNumber; }

        /// <summary>SUMIFS：主体 + 客商 + 年度 +（可选）条件列。</summary>
        internal string Sumifs(int rowNumber, string sumLetter, Letters letters,
            MatrixColumn debitColumn, int firstHeaderRow)
        {
            StringBuilder sb = new StringBuilder("=SUMIFS(");
            sb.Append(Col(sumLetter));
            sb.Append(",").Append(Col(letters.Entity)).Append(",").Append(EntityRef(rowNumber));
            sb.Append(",").Append(Col(letters.Counterparty)).Append(",").Append(CustomerRef(rowNumber));
            sb.Append(",").Append(Col(letters.Year)).Append(",")
              .Append(debitColumn.Letter).Append("$").Append(firstHeaderRow);
            if (options.HasConditionColumn && letters.Condition != null)
                sb.Append(",").Append(Col(letters.Condition)).Append(",").Append(Criterion(options.ConditionValue));
            sb.Append(")");
            return sb.ToString();
        }

        /// <summary>
        /// SUMPRODUCT：现算分支专用。**强制限定数据区（$O$3:$O$末行），忽略「引用整列」**——
        /// 整列会把第 1/2 行的汇总行与表头文本卷入算术，文本参与乘法产生 #VALUE!；
        /// WPS 表格对 SUMPRODUCT 内 IFERROR 的数组语义支持不完整，实测无法拦住这种传染
        /// （2026-10-16 实机验收 K11 = #VALUE!）。区间从数据首行起、不含文本，天然无错误值，
        /// 同时把扫描行数从 104 万降到实际数据行数，重算性能也一并解决。
        /// </summary>
        internal string Sumproduct(int rowNumber, Letters letters,
            MatrixColumn debitColumn, int firstHeaderRow, string[] factors)
        {
            // 条件数组必须与金额因子用**同一个区间**：SUMPRODUCT 要求所有相乘数组维度一致，
            // 一旦条件用整列（1048576 行）而因子用区间（2693 行）就会返回 #N/A（实机验收 2026-10-16 第二例）。
            StringBuilder sb = new StringBuilder("=SUMPRODUCT(");
            sb.Append("(").Append(ColData(letters.Entity)).Append("=").Append(EntityRef(rowNumber)).Append(")*");
            sb.Append("(").Append(ColData(letters.Counterparty)).Append("=").Append(CustomerRef(rowNumber)).Append(")*");
            sb.Append("(").Append(ColData(letters.Year)).Append("=")
              .Append(debitColumn.Letter).Append("$").Append(firstHeaderRow).Append(")*");
            if (options.HasConditionColumn && letters.Condition != null)
                sb.Append("(").Append(ColData(letters.Condition)).Append("=")
                  .Append(Criterion(options.ConditionValue)).Append(")*");
            for (int i = 0; i < factors.Length; i++)
            {
                EliminationCancellation.Check();
                if (i > 0) sb.Append("*");
                sb.Append("IFERROR(").Append(factors[i]).Append(",0)");
            }
            sb.Append(")");
            return sb.ToString();
        }
    }
}
