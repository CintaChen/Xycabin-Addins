using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
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

    /// <summary>矩阵底稿构建器（纯核，零 COM 零 WinForms）。</summary>
    public static partial class EliminationMatrixBuilder
    {
        /// <summary>固定信息列的第 2 行表头。</summary>


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
}
