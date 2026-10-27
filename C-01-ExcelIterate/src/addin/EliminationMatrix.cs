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

    /// <summary>行模板登记表。第二阶段接入往来抵消只需在此新增一个方法。</summary>
    public static class MatrixTemplates
    {
        public static IList<MatrixRowTemplate> RevenueCost()
        {
            List<MatrixRowTemplate> list = new List<MatrixRowTemplate>();

            MatrixRowTemplate revenue = new MatrixRowTemplate();
            revenue.Key = "Revenue";
            revenue.Account = EliminationEngine.AccRevenue;
            revenue.IsDebit = true;
            revenue.Adjustment = "收入成本抵消";
            revenue.NoteTemplate = "{Entity} → {Customer}：收入抵消";
            revenue.Terms.Add(new MatrixTerm { Plus = true, Kind = MatrixCellKind.Measure, Key = "Revenue" });
            list.Add(revenue);

            MatrixRowTemplate cost = new MatrixRowTemplate();
            cost.Key = "Cost";
            cost.Account = EliminationEngine.AccCost;
            cost.IsDebit = false;
            cost.Adjustment = "收入成本抵消";
            cost.NoteTemplate = "{Entity} → {Customer}：成本抵消";
            cost.Terms.Add(new MatrixTerm { Plus = true, Kind = MatrixCellKind.Measure, Key = "Cost" });
            list.Add(cost);

            MatrixRowTemplate restore = new MatrixRowTemplate();
            restore.Key = "Restore";
            restore.Account = EliminationEngine.AccCost;
            restore.IsDebit = false;
            restore.Adjustment = "收入成本抵消";
            restore.NoteTemplate = "{Customer}：成本还原";
            // （收入 − 成本）− 未实现损益：明确分组，与底稿习惯一致
            restore.ExpressionTemplate = "=({0}-{1})-{2}";
            restore.Terms.Add(new MatrixTerm { Plus = true, Kind = MatrixCellKind.RowRef, Key = "Revenue" });
            restore.Terms.Add(new MatrixTerm { Plus = false, Kind = MatrixCellKind.RowRef, Key = "Cost" });
            restore.Terms.Add(new MatrixTerm { Plus = false, Kind = MatrixCellKind.RowRef, Key = "Unrealized" });
            list.Add(restore);

            MatrixRowTemplate unrealized = new MatrixRowTemplate();
            unrealized.Key = "Unrealized";
            unrealized.Account = EliminationEngine.AccInventory;
            unrealized.IsDebit = false;
            unrealized.Adjustment = "收入成本抵消";
            unrealized.NoteTemplate = "{Entity} → {Customer}：未实现内部交易损益";
            unrealized.Terms.Add(new MatrixTerm { Plus = true, Kind = MatrixCellKind.Measure, Key = "Unrealized" });
            list.Add(unrealized);

            return list;
        }
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

    /// <summary>一组（核算主体, 客商）在某一年度的聚合值。</summary>
    internal sealed class MatrixMeasure
    {
        public decimal Revenue;
        public decimal Cost;
        /// <summary>源表「未实现内部交易损益」列有可解析数值的行数。</summary>
        public int ProfitColumnRows;
        public decimal ProfitColumnSum;
        /// <summary>逐行 Σ(收入−成本)×未对外销售比例。</summary>
        public decimal ComputedProfit;
        public int MatchedRows;

        internal bool UseProfitColumn { get { return ProfitColumnRows > 0; } }

        internal decimal Unrealized()
        { return UseProfitColumn ? ProfitColumnSum : ComputedProfit; }
    }

    /// <summary>矩阵底稿构建器（纯核，零 COM 零 WinForms）。</summary>
    public static class EliminationMatrixBuilder
    {
        /// <summary>固定信息列的第 2 行表头。</summary>
        public static readonly string[] FixedHeaders = new string[]
        { "核算主体", "明细【客商/二级明细】", "备注", "调整事项", "会计科目" };

        /// <summary>固定信息列数。</summary>
        public const int FixedColumns = 5;

        /// <summary>
        /// 主构建流程：筛选 → 三维聚合 → 年度降序 → 套行模板 → 生成金额（值 + 公式）。
        /// </summary>
        public static EliminationMatrix Build(IList<EliminationRow> rows,
            IList<EliminationColumn> columns, MatrixOptions options)
        {
            EliminationMatrix matrix = new EliminationMatrix();
            if (rows == null || columns == null) return matrix;
            if (options == null) options = new MatrixOptions();

            Letters letters = new Letters(columns);

            List<string> order = new List<string>();
            Dictionary<string, Dictionary<string, MatrixMeasure>> groups =
                new Dictionary<string, Dictionary<string, MatrixMeasure>>();
            List<string> years = new List<string>();

            Filter(rows, options, letters, matrix, groups, order, years);

            if (!options.HasConditionColumn)
                AddWarning(matrix, "未指定「参与抵消条件列」，全部数据行参与计算。");
            if (letters.Ratio == null && letters.Profit == null)
                matrix.Errors.Add("未映射「未实现内部交易损益」列，也未映射「未对外销售比例」列，无法计算未实现损益。");

            years.Sort(CompareYearDesc);
            matrix.Years = years;
            if (years.Count > options.MaxYears)
            {
                matrix.Errors.Add("年度数量 " + years.Count + " 超过上限 " + options.MaxYears
                    + "，请确认是否选错了年份列。");
                return matrix;
            }
            if (years.Count == 0)
            {
                matrix.Errors.Add("没有符合条件的数据行，无法生成底稿。");
                return matrix;
            }
            matrix.Columns = BuildColumns(years, options.StartColumn);

            Emit(matrix, groups, order, years, options, letters);
            VerifyIdentity(matrix);
            return matrix;
        }

        // ---------------- ① 筛选 + ② 三维聚合 ----------------

        private static void Filter(IList<EliminationRow> rows, MatrixOptions options, Letters letters,
            EliminationMatrix matrix, Dictionary<string, Dictionary<string, MatrixMeasure>> groups,
            List<string> order, List<string> years)
        {
            string want = options.ConditionValue == null ? "" : options.ConditionValue.Trim();
            for (int i = 0; i < rows.Count; i++)
            {
                EliminationCancellation.Check();
                EliminationRow row = rows[i];
                if (options.HasConditionColumn)
                {
                    string actual = Text(row, EliminationRole.Condition).Trim();
                    if (!string.Equals(actual, want, StringComparison.OrdinalIgnoreCase))
                    {
                        matrix.FilteredRows++;
                        continue;
                    }
                }

                string entity = Text(row, EliminationRole.Entity).Trim();
                string customer = Text(row, EliminationRole.Counterparty).Trim();
                string year = ParseYear(Text(row, EliminationRole.Year));
                if (entity.Length == 0)
                {
                    matrix.SkippedRows++;
                    AddWarning(matrix, "第 " + row.SourceRow + " 行缺少核算主体，已跳过。");
                    continue;
                }
                if (customer.Length == 0)
                {
                    matrix.SkippedRows++;
                    AddWarning(matrix, "第 " + row.SourceRow + " 行缺少客户名称，已跳过。");
                    continue;
                }
                if (year == null)
                {
                    matrix.SkippedRows++;
                    AddWarning(matrix, "第 " + row.SourceRow + " 行年份无法识别（需形如 2024），已跳过。");
                    continue;
                }
                decimal revenue, cost;
                if (!EliminationEngine.TryParseAmount(Text(row, EliminationRole.Revenue), out revenue)
                    || !EliminationEngine.TryParseAmount(Text(row, EliminationRole.Cost), out cost))
                {
                    matrix.SkippedRows++;
                    AddWarning(matrix, "第 " + row.SourceRow + " 行收入或成本无法解析，已跳过。");
                    continue;
                }

                string key = entity + "\u0001" + customer;
                Dictionary<string, MatrixMeasure> byYear;
                if (!groups.TryGetValue(key, out byYear))
                {
                    byYear = new Dictionary<string, MatrixMeasure>();
                    groups[key] = byYear;
                    order.Add(key);
                }
                MatrixMeasure measure;
                if (!byYear.TryGetValue(year, out measure))
                {
                    measure = new MatrixMeasure();
                    byYear[year] = measure;
                    if (!years.Contains(year)) years.Add(year);
                }
                measure.Revenue += revenue;
                measure.Cost += cost;
                measure.MatchedRows++;

                decimal profit;
                if (letters.Profit != null
                    && EliminationEngine.TryParseAmount(Text(row, EliminationRole.UnrealizedProfit), out profit))
                {
                    measure.ProfitColumnRows++;
                    measure.ProfitColumnSum += profit;
                }
                decimal ratio;
                if (letters.Ratio != null
                    && EliminationEngine.TryParseRatio(Text(row, EliminationRole.UnsoldRatio), out ratio))
                {
                    // 必须逐行加权：比例可能逐行不同，先汇总再乘会失真。
                    measure.ComputedProfit += (revenue - cost) * ratio;
                }
                if (letters.Profit == null)
                    AddWarning(matrix, "未映射「未实现内部交易损益」列，未实现损益按（收入−成本）×未对外销售比例现算。");
                matrix.MatchedRows++;
            }
        }

        // ---------------- ④ 套行模板 ----------------

        private static void Emit(EliminationMatrix matrix,
            Dictionary<string, Dictionary<string, MatrixMeasure>> groups,
            List<string> order, List<string> years, MatrixOptions options, Letters letters)
        {
            IList<MatrixRowTemplate> templates = MatrixTemplates.RevenueCost();
            int height = templates.Count;
            matrix.TemplateCount = height;
            options.GroupHeight = height;
            int firstHeaderRow = options.StartRow + (options.TotalRow ? 1 : 0);
            int dataTop = firstHeaderRow + 2;
            int stride = height + (options.GroupGap > 0 ? options.GroupGap : 0);
            // 存回结果：Balanced 与 VerifyIdentity 都要按同一个跨行数推导组首行，三处公式必须一致
            matrix.GroupGap = options.GroupGap;
            matrix.GroupStride = stride;

            for (int g = 0; g < order.Count; g++)
            {
                EliminationCancellation.Check();
                string key = order[g];
                string[] parts = key.Split('\u0001');
                string entity = parts[0];
                string customer = parts.Length > 1 ? parts[1] : "";
                Dictionary<string, MatrixMeasure> byYear = groups[key];
                // 组间距：每个交易对之间空 GroupGap 行（组内 RowRef 引用 Top+index 自动跟随）
                int top = dataTop + g * stride;

                // 记忆化：行模板之间可互相引用（成本还原引用收入/成本/存货）。
                // key 必须含年度——同一个模板行在不同年度的值不同，按模板索引缓存会串。
                Dictionary<string, MatrixAmount> memo = new Dictionary<string, MatrixAmount>();
                for (int t = 0; t < height; t++)
                {
                    EliminationCancellation.Check();
                    MatrixRowTemplate template = templates[t];
                    MatrixRow outRow = new MatrixRow();
                    outRow.RowNumber = top + t;
                    outRow.Entity = entity;
                    outRow.Customer = customer;
                    outRow.Account = template.Account;
                    outRow.Adjustment = template.Adjustment;
                    outRow.TemplateKey = template.Key;
                    outRow.Note = template.NoteTemplate
                        .Replace("{Entity}", entity).Replace("{Customer}", customer);

                    for (int y = 0; y < years.Count; y++)
                    {
                        EliminationCancellation.Check();
                        string year = years[y];
                        MatrixMeasure measure = byYear.ContainsKey(year) ? byYear[year] : null;
                        // 该交易对在本年度没有数据 → 整列留空，不写返回 0 的冗余公式
                        if (measure == null) continue;
                        MatrixColumn debitColumn = matrix.DebitColumn(year);
                        MatrixColumn creditColumn = matrix.CreditColumn(year);
                        MatrixColumn target = template.IsDebit ? debitColumn : creditColumn;
                        if (target == null) continue;

                        GroupScope scope = new GroupScope();
                        scope.Measure = measure;
                        scope.Top = top;
                        scope.RowNumber = outRow.RowNumber;
                        scope.Templates = templates;
                        scope.DebitColumn = debitColumn;
                        scope.CreditColumn = creditColumn;
                        scope.FirstHeaderRow = firstHeaderRow;
                        scope.Options = options;
                        scope.Letters = letters;
                        scope.Memo = memo;
                        scope.Year = year;
                        scope.Depth = 0;

                        MatrixAmount amount = Resolve(t, scope);
                        if (amount == null) continue;
                        memo[t + "\u0001" + year] = amount;
                        if (template.IsDebit) outRow.Debit[year] = amount;
                        else outRow.Credit[year] = amount;
                    }
                    matrix.Rows.Add(outRow);
                }
            }
            matrix.GroupCount = order.Count;
        }

        /// <summary>递归求值一行模板（RowRef 会回溯被引用行，带记忆化与深度保护）。</summary>
        private static MatrixAmount Resolve(int templateIndex, GroupScope scope)
        {
            MatrixAmount cached;
            if (scope.Memo.TryGetValue(templateIndex + "\u0001" + scope.Year, out cached)) return cached;

            MatrixRowTemplate template = scope.Templates[templateIndex];
            StringBuilder formula = new StringBuilder();
            decimal value = 0m;
            bool any = false;
            List<string> parts = new List<string>();

            for (int i = 0; i < template.Terms.Count; i++)
            {
                EliminationCancellation.Check();
                MatrixTerm term = template.Terms[i];
                decimal amount;
                string text;

                if (term.Kind == MatrixCellKind.Measure)
                {
                    amount = MeasureValue(term.Key, scope.Measure);
                    text = MeasureFormula(term.Key, scope);
                }
                else
                {
                    int index = IndexOfTemplate(scope.Templates, term.Key);
                    if (index < 0 || scope.Depth > 8) continue;
                    MatrixRowTemplate referenced = scope.Templates[index];
                    MatrixColumn refColumn = referenced.IsDebit ? scope.DebitColumn : scope.CreditColumn;
                    if (refColumn == null) continue;

                    GroupScope deeper = scope.Clone();
                    deeper.Depth = scope.Depth + 1;
                    deeper.RowNumber = scope.Top + index;
                    MatrixAmount referencedAmount = Resolve(index, deeper);
                    if (referencedAmount == null) continue;
                    amount = referencedAmount.Value;
                    text = refColumn.Letter + (scope.Top + index);
                }

                if (text == null) continue;
                // 统一剥掉各项自带的前导 "="，最后由本方法决定整个公式的前缀，
                // 否则会出现 "==SUMIFS(" 或 RowRef 组合时整体缺 "=" 而被误判成文本。
                if (text.StartsWith("=", StringComparison.Ordinal)) text = text.Substring(1);

                value += term.Plus ? amount : -amount;
                parts.Add(text);
                if (!any)
                {
                    formula.Append(term.Plus ? "" : "-");
                    formula.Append(Wrap(text));
                }
                else
                {
                    formula.Append(term.Plus ? "+" : "-");
                    formula.Append(Wrap(text));
                }
                any = true;
            }

            if (!any) return null;
            MatrixAmount result = new MatrixAmount();
            result.Value = Money2(value);
            if (template.ExpressionTemplate != null && template.ExpressionTemplate.Length > 0)
            {
                // 有表达式模板：按 Terms 顺序代入占位符（模板自带前导 "="）
                StringBuilder composed = new StringBuilder(template.ExpressionTemplate);
                for (int i = 0; i < parts.Count; i++)
                    composed.Replace("{" + i + "}", parts[i]);
                result.Formula = composed.ToString();
            }
            else
            {
                result.Formula = "=" + formula.ToString();
            }
            return result;
        }

        private static string Wrap(string text)
        {
            if (text == null) return "";
            if (IsFunctionCall(text)) return text;
            if (text.IndexOf('-') > 0 || text.IndexOf('+') > 0) return "(" + text + ")";
            return text;
        }

        /// <summary>
        /// 是否整体是一个函数调用（形如 SUMIFS(...) / SUMPRODUCT(...)，已剥掉前导 "="）。
        /// 是则不能再包括号——否则会写成 =(SUMIFS(...)) 这种冗余形式。
        /// </summary>
        private static bool IsFunctionCall(string text)
        {
            if (text == null || text.Length < 3) return false;
            int i = 0;
            while (i < text.Length && char.IsLetter(text[i])) i++;
            return i > 0 && i < text.Length && text[i] == '(';
        }

        private static decimal MeasureValue(string key, MatrixMeasure measure)
        {
            if (measure == null) return 0m;
            if (key == "Revenue") return Money2(measure.Revenue);
            if (key == "Cost") return Money2(measure.Cost);
            if (key == "Unrealized") return Money2(measure.Unrealized());
            return 0m;
        }

        /// <summary>度量键 → 公式文本。Unrealized 在损益列不可用时改用 SUMPRODUCT 现算。</summary>
        private static string MeasureFormula(string key, GroupScope scope)
        {
            Letters letters = scope.Letters;
            string sumLetter;
            if (key == "Revenue") sumLetter = letters.Revenue;
            else if (key == "Cost") sumLetter = letters.Cost;
            else if (key == "Unrealized") sumLetter = letters.Profit;
            else return null;

            FormulaKit kit = new FormulaKit(scope.Options);
            // 损益列未映射或本组该列无值 → 改用 (收入−成本)×比例 现算（只能 SUMPRODUCT）。
            // 必须先判这个分支：否则 sumLetter 为 null 会提前返回，永远走不到现算。
            if (key == "Unrealized" && !(scope.Measure != null && scope.Measure.UseProfitColumn))
            {
                if (letters.Revenue == null || letters.Cost == null || letters.Ratio == null) return null;
                // 因子必须用数据区区间（ColData）：整列会卷入表头/汇总行文本，#VALUE! 传染整个 SUMPRODUCT
                return kit.Sumproduct(scope.RowNumber, letters, scope.DebitColumn, scope.FirstHeaderRow,
                    new string[] { kit.ColData(letters.Revenue) + "-" + kit.ColData(letters.Cost), kit.ColData(letters.Ratio) });
            }
            if (sumLetter == null) return null;
            return kit.Sumifs(scope.RowNumber, sumLetter, letters, scope.DebitColumn, scope.FirstHeaderRow);
        }

        private static int IndexOfTemplate(IList<MatrixRowTemplate> templates, string key)
        {
            for (int i = 0; i < templates.Count; i++)
                if (templates[i].Key == key) return i;
            return -1;
        }

        /// <summary>恒等式校验：每组每年 Σ借 == Σ贷。失败即代码缺陷，写 Errors 阻断。</summary>
        private static void VerifyIdentity(EliminationMatrix matrix)
        {
            if (matrix.GroupCount == 0 || matrix.Rows.Count == 0 || matrix.TemplateCount == 0) return;
            int dataTop = matrix.Rows[0].RowNumber;
            int stride = matrix.GroupStride > 0 ? matrix.GroupStride : matrix.TemplateCount;
            for (int g = 0; g < matrix.GroupCount; g++)
            {
                EliminationCancellation.Check();
                int top = dataTop + g * stride;
                for (int y = 0; y < matrix.Years.Count; y++)
                {
                    EliminationCancellation.Check();
                    string year = matrix.Years[y];
                    decimal debit = 0m, credit = 0m;
                    for (int r = 0; r < matrix.Rows.Count; r++)
                    {
                        EliminationCancellation.Check();
                        MatrixRow row = matrix.Rows[r];
                        if (row.RowNumber < top || row.RowNumber >= top + matrix.TemplateCount) continue;
                        MatrixAmount amount;
                        if (row.Debit.TryGetValue(year, out amount)) debit += amount.Value;
                        if (row.Credit.TryGetValue(year, out amount)) credit += amount.Value;
                    }
                    if (Math.Abs(debit - credit) > 0.005m)
                    {
                        matrix.Errors.Add("恒等式校验失败：第 " + top + " 行起这组在 " + year + " 年借 "
                            + debit.ToString("0.00", CultureInfo.InvariantCulture) + " / 贷 "
                            + credit.ToString("0.00", CultureInfo.InvariantCulture) + "，请检查。");
                    }
                }
            }
        }

        // ---------------- 列布局 ----------------

        /// <summary>固定五列 + 每年度三列（[分隔][借][贷]）。</summary>
        public static List<MatrixColumn> BuildColumns(IList<string> years, int startColumn)
        {
            List<MatrixColumn> list = new List<MatrixColumn>();
            for (int i = 0; i < FixedColumns; i++)
            {
                EliminationCancellation.Check();
                MatrixColumn column = new MatrixColumn();
                column.Index = i;
                column.Number = startColumn + i;
                column.Letter = Letter(column.Number);
                column.Head2 = FixedHeaders[i];
                list.Add(column);
            }
            for (int i = 0; i < years.Count; i++)
            {
                EliminationCancellation.Check();
                int spacerIndex = FixedColumns + i * 3;
                MatrixColumn spacer = new MatrixColumn();
                spacer.Index = spacerIndex;
                spacer.Number = startColumn + spacerIndex;
                spacer.Letter = Letter(spacer.Number);
                spacer.IsSpacer = true;
                list.Add(spacer);

                int debitIndex = spacerIndex + 1;
                MatrixColumn debit = new MatrixColumn();
                debit.Index = debitIndex;
                debit.Number = startColumn + debitIndex;
                debit.Letter = Letter(debit.Number);
                debit.Year = years[i];
                debit.IsDebit = true;
                debit.Head2 = years[i];
                debit.Head3 = "借方";
                list.Add(debit);

                int creditIndex = spacerIndex + 2;
                MatrixColumn credit = new MatrixColumn();
                credit.Index = creditIndex;
                credit.Number = startColumn + creditIndex;
                credit.Letter = Letter(credit.Number);
                credit.Year = years[i];
                credit.IsDebit = false;
                credit.Head2 = years[i];
                credit.Head3 = "贷方";
                list.Add(credit);
            }
            return list;
        }

        /// <summary>
        /// 从年份文本提取四位年份。只认 (19|20)\d{2}，不回落数值解析——
        /// 否则 Excel 日期序列号（如 45658）会被误当成年份。
        /// </summary>
        public static string ParseYear(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            Match match = Regex.Match(text.Trim(), "(19|20)\\d{2}");
            return match.Success ? match.Value : null;
        }

        private static int CompareYearDesc(string a, string b)
        {
            int ia, ib;
            bool oa = int.TryParse(a, NumberStyles.Integer, CultureInfo.InvariantCulture, out ia);
            bool ob = int.TryParse(b, NumberStyles.Integer, CultureInfo.InvariantCulture, out ib);
            if (oa && ob) return ib.CompareTo(ia);
            if (oa) return -1;
            if (ob) return 1;
            return string.Compare(b, a, StringComparison.Ordinal);
        }

        /// <summary>列号 → 列字母（1 → "A"，27 → "AA"）。</summary>
        public static string Letter(int column)
        {
            if (column <= 0) return "A";
            StringBuilder sb = new StringBuilder();
            int rest = column;
            while (rest > 0)
            {
                int mod = (rest - 1) % 26;
                sb.Insert(0, (char)('A' + mod));
                rest = (rest - mod - 1) / 26;
            }
            return sb.ToString();
        }

        private static string Text(EliminationRow row, EliminationRole role)
        {
            string value;
            return row.Values.TryGetValue(role, out value) ? (value ?? "").Trim() : "";
        }

        private static void AddWarning(EliminationMatrix matrix, string text)
        {
            if (matrix.Warnings.Contains(text)) return;
            if (matrix.Warnings.Count < 200) matrix.Warnings.Add(text);
        }

        private static decimal Money2(decimal value)
        { return Math.Round(value, 2, MidpointRounding.AwayFromZero); }
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

    /// <summary>递归求值时的一组上下文（同一交易对的同一年度）。</summary>
    internal sealed class GroupScope
    {
        internal MatrixMeasure Measure;
        internal int Top;
        internal int RowNumber;
        internal IList<MatrixRowTemplate> Templates;
        internal MatrixColumn DebitColumn;
        internal MatrixColumn CreditColumn;
        internal int FirstHeaderRow;
        internal MatrixOptions Options;
        internal Letters Letters;
        internal Dictionary<string, MatrixAmount> Memo;
        /// <summary>当前年度；记忆化 key 的组成部分，递归时保持不变。</summary>
        internal string Year = "";
        internal int Depth;

        internal GroupScope Clone()
        {
            GroupScope copy = new GroupScope();
            copy.Measure = Measure;
            copy.Top = Top;
            copy.RowNumber = RowNumber;
            copy.Templates = Templates;
            copy.DebitColumn = DebitColumn;
            copy.CreditColumn = CreditColumn;
            copy.FirstHeaderRow = FirstHeaderRow;
            copy.Options = Options;
            copy.Letters = Letters;
            copy.Memo = Memo;
            copy.Year = Year;
            copy.Depth = Depth;
            return copy;
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
