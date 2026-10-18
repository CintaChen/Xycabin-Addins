using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Excel_Iterate
{
    // =====================================================================
    // 往来抵消 —— 两张输出表的纯核（零 COM 零 WinForms，探针可注入假数据）
    //
    // 源表「往来明细表」：用户选列 → 截止日（取前 4 位作年度）、核算主体、
    // 科目、对方单位、是否合并内关联方、审定数。
    //
    // 产出两张表：
    //   ① 往来抵消差异：数据驱动的公司对 × 年度 × 3 个科目对，
    //      按「净额法」取数：资产端 = 债权方对债务方的资产净额（资产−负债），
    //      负债端 = 债务方对债权方的负债净额（负债−资产），
    //      抵消=MIN(资产端,负债端)，差额=资产端−负债端，并按净额关系写备注。
    //   ② 往来抵消结果：与「收入抵消结果」同构的矩阵式底稿，
    //      每个有金额科目对 2 行（债权方借资产科目 / 债务方贷负债科目），
    //      仅当抵消金额不为 0 时才输出；同一（主体,对方,科目）跨多年度
    //      合并为一行，金额分列到各年度列。
    //
    // 取数口径（净额法）：
    //   资产端账面净额 = 债权方互指债务方、科目=资产科目 的审定数 − 科目=负债科目 的审定数
    //   负债端账面净额 = 债务方互指债权方、科目=负债科目 的审定数 − 科目=资产科目 的审定数
    //   债权方 = 净额(资产−负债)较大的一方（动态判定），两列恒非负。
    //
    // 设计取向：与收入成本的矩阵核（EliminationMatrix.cs）并列而非交织——
    // 往来抵消的「组内行数随科目对数量变化」，若塞进固定模板机制会破坏
    // TemplateCount/GroupStride 的既有假设。这里复用 MatrixColumn/MatrixRow/
    // MatrixAmount 等类型与 Host 的写入链路，但构建逻辑独立。
    // =====================================================================

    /// <summary>一个科目对（资产科目 ↔ 负债科目）。</summary>
    public sealed class DiffAccountPair
    {
        public string Asset = "";
        public string Liability = "";
        /// <summary>源表可能的别名（源表科目名与输出名不一致时用）。</summary>
        public string AssetAlias = "";
        public string LiabilityAlias = "";

        internal DiffAccountPair(string asset, string liability, string assetAlias, string liabilityAlias)
        {
            Asset = asset; Liability = liability;
            AssetAlias = assetAlias; LiabilityAlias = liabilityAlias;
        }
    }

    /// <summary>差异表的一行。</summary>
    public sealed class EliminationDiffRow
    {
        public int Index;
        public string Year = "";
        /// <summary>核算主体（债权方）。</summary>
        public string Creditor = "";
        /// <summary>对方单位（债务方）。</summary>
        public string Debtor = "";
        public string AssetAccount = "";
        public string LiabilityAccount = "";
        public MatrixAmount Asset;      // 资产端账面净额(审定数)
        public MatrixAmount Liability;  // 负债端账面净额(审定数)
        public MatrixAmount Offset;     // 抵消金额(以小抵大)
        public MatrixAmount Difference; // 差额(资产-负债)
        /// <summary>坏账准备抵消额：债权方被抵消的资产科目对应的坏账准备（全额，AM 列）。</summary>
        public decimal BadDebt;
        public string Note = "";
        /// <summary>
        /// 供「结果表」使用的跨表引用公式：='差异表'!I&lt;本行号&gt;（I 列即差异表的抵消金额列）。
        /// 结果表不能复用 Offset.Formula——那是差异表内部的 =MIN(G?,H?)，
        /// 而结果表的 G/H 是年度金额列、行号也不同，复用会引用错乱（实机验收发现）。
        /// </summary>
        public string OffsetReference;
        public int RowNumber;
        /// <summary>所属主体对序号（空行粒度用）。</summary>
        public int GroupIndex;
    }

    /// <summary>「往来抵消差异」表。</summary>
    public sealed class EliminationDiffTable
    {
        public static readonly string[] Headers = new string[]
        {
            "序号", "年度", "核算主体(债权方)", "对方单位(债务方)", "资产科目", "负债科目",
            "资产端账面净额(审定数)", "负债端账面净额(审定数)", "抵消金额(以小抵大)",
            "差额(资产-负债)", "备注"
        };

        public List<string> Years = new List<string>();
        public List<EliminationDiffRow> Rows = new List<EliminationDiffRow>();
        public List<string> Warnings = new List<string>();
        public List<string> Errors = new List<string>();
        public int GroupCount;
        /// <summary>每个 (主体对 × 年度) 的行数 = 科目对数量。</summary>
        public int RowsPerYear = 3;
        public int GapRows;
        public int MatchedRows;
        public int FilteredRows;
        public int SkippedRows;

        public int ColumnCount { get { return Headers.Length; } }
        public int RowCount { get { return Rows.Count; } }
    }

    /// <summary>往来抵消构建器。</summary>
    public static class EliminationDiffBuilder
    {
        /// <summary>坏账准备抵消的借/贷科目与占位符。</summary>
        internal const string BadDebtAccountName = "坏账准备";
        internal const string BadDebtMgmtAccountName = "管理费用";
        internal const string BadDebtMgmtCustomer = "—";      // 贷管理费用无对方单位，B 列占位符
        internal const string BadDebtAdjustment = "坏账准备抵消";

        /// <summary>三个科目对（顺序即行序）。</summary>
        public static IList<DiffAccountPair> Pairs()
        {
            List<DiffAccountPair> list = new List<DiffAccountPair>();
            list.Add(new DiffAccountPair("应收账款", "应付账款", "", ""));
            list.Add(new DiffAccountPair("其他应收款", "其他应付款", "", ""));
            // 源表实际科目为「预付款项 / 预收款项」，输出按样张名写「预付账款 / 合同负债」
            list.Add(new DiffAccountPair("预付账款", "合同负债", "预付款项", "预收款项"));
            return list;
        }

        /// <summary>
        /// 主构建：筛条件列 → 数据驱动收集公司对 + 归集金额 → 净额法
        /// （每个公司对 × 年度 × 科目对：双向四口径 → 净额 → 动态债权方 → 抵消/差额/备注）。
        /// </summary>
        public static void Build(IList<EliminationRow> rows, IList<EliminationColumn> columns,
            MatrixOptions options, out EliminationDiffTable diff, out EliminationMatrix result)
        {
            diff = new EliminationDiffTable();
            result = new EliminationMatrix();
            if (rows == null || columns == null) return;
            if (options == null) options = new MatrixOptions();

            Letters letters = new Letters(columns);
            string want = options.ConditionValue == null ? "" : options.ConditionValue.Trim();

            // ---- ① 筛选 + 数据驱动收集公司对 + 归集金额 ----
            // key = 主体|对方(规范名)|年度|科目 → 金额
            Dictionary<string, decimal> sums = new Dictionary<string, decimal>(StringComparer.Ordinal);
            // key = 主体|对方(规范名)|年度|科目 → 坏账准备（AM 列，可选映射）
            Dictionary<string, decimal> badDebtSums = new Dictionary<string, decimal>(StringComparer.Ordinal);
            bool hasBadDebtColumn = EliminationEngine.HasRole(columns, EliminationRole.BadDebt);
            List<string> years = new List<string>();
            HashSet<string> accountValues = new HashSet<string>(StringComparer.Ordinal);
            // 数据驱动配对：真实出现过的无序公司对（规范名，排序去重）
            HashSet<string> pairSet = new HashSet<string>(StringComparer.Ordinal);

            for (int i = 0; i < rows.Count; i++)
            {
                EliminationRow row = rows[i];
                if (options.HasConditionColumn)
                {
                    string actual = Text(row, EliminationRole.Condition).Trim();
                    if (!string.Equals(actual, want, StringComparison.OrdinalIgnoreCase))
                    { diff.FilteredRows++; continue; }
                }
                string entity = CanonicalName(Text(row, EliminationRole.Entity));
                string party = CanonicalName(Text(row, EliminationRole.Counterparty));
                string account = Text(row, EliminationRole.Account).Trim();
                string year = EliminationMatrixBuilder.ParseYear(Text(row, EliminationRole.Deadline));
                decimal amount;
                if (!EliminationEngine.TryParseAmount(Text(row, EliminationRole.Amount), out amount))
                    amount = 0m;

                if (entity.Length == 0 || year == null)
                { diff.SkippedRows++; continue; }
                if (!years.Contains(year)) years.Add(year);
                if (account.Length > 0 && !accountValues.Contains(account)) accountValues.Add(account);

                // 数据驱动配对：只记真实存在、且非自配的无序公司对
                if (party.Length > 0 && entity != party)
                {
                    string pk = string.Compare(entity, party, StringComparison.Ordinal) <= 0
                        ? entity + "\u0001" + party
                        : party + "\u0001" + entity;
                    if (!pairSet.Contains(pk)) pairSet.Add(pk);
                }

                string key = entity + "\u0001" + party + "\u0001" + year + "\u0001" + account;
                decimal prior;
                sums.TryGetValue(key, out prior);
                sums[key] = prior + amount;

                // 坏账准备（AM 列，可选映射；仅当该行有坏账准备时归集）
                if (hasBadDebtColumn)
                {
                    decimal badDebt;
                    if (EliminationEngine.TryParseAmount(Text(row, EliminationRole.BadDebt), out badDebt) && badDebt != 0m)
                    {
                        decimal bdPrior;
                        badDebtSums.TryGetValue(key, out bdPrior);
                        badDebtSums[key] = bdPrior + badDebt;
                    }
                }
                diff.MatchedRows++;
            }

            if (pairSet.Count == 0)
            {
                diff.Errors.Add("符合条件的核算主体不足 2 个，无法生成往来抵消表。");
                return;
            }
            years.Sort(CompareYearDesc);
            diff.Years = years;
            result.Years = years;

            // ---- ② 公司对排序（按规范名）后，逐对 × 年度 × 科目对 → 净额法 ----
            IList<DiffAccountPair> pairs = Pairs();
            diff.RowsPerYear = pairs.Count;
            diff.GapRows = options.GroupGap;
            List<string[]> orderedPairs = new List<string[]>();
            foreach (string pk in pairSet) orderedPairs.Add(pk.Split('\u0001'));
            orderedPairs.Sort(delegate(string[] x, string[] y)
            {
                int c = string.Compare(x[0], y[0], StringComparison.Ordinal);
                return c != 0 ? c : string.Compare(x[1], y[1], StringComparison.Ordinal);
            });

            int groupIndex = 0;
            int rowNumber = 1;   // 差异表第 1 行是表头，数据从第 2 行起
            int sequence = 0;

            foreach (string[] pair in orderedPairs)
            {
                string first = pair[0];
                string second = pair[1];
                for (int y = 0; y < years.Count; y++)
                {
                    for (int p = 0; p < pairs.Count; p++)
                    {
                        DiffAccountPair pairAcc = pairs[p];
                        string assetAcc = SourceAccount(accountValues, pairAcc.Asset, pairAcc.AssetAlias);
                        string liabAcc = SourceAccount(accountValues, pairAcc.Liability, pairAcc.LiabilityAlias);

                        // 双向四口径（精确匹配）
                        decimal aAsset = DirectionalSum(sums, first, second, years[y], assetAcc);
                        decimal aLiab = DirectionalSum(sums, first, second, years[y], liabAcc);
                        decimal bAsset = DirectionalSum(sums, second, first, years[y], assetAcc);
                        decimal bLiab = DirectionalSum(sums, second, first, years[y], liabAcc);

                        decimal netA = aAsset - aLiab;   // 第一公司净额（>0 = 净债权方）
                        decimal netB = bAsset - bLiab;   // 第二公司净额

                        // 动态判定债权方/债务方（两列恒非负）
                        string creditor, debtor;
                        decimal assetNet, liabilityNet;
                        if (netA >= netB)
                        {
                            creditor = first; debtor = second;
                            assetNet = netA; liabilityNet = -netB;
                        }
                        else
                        {
                            creditor = second; debtor = first;
                            assetNet = netB; liabilityNet = -netA;
                        }
                        if (assetNet < 0m) assetNet = 0m;
                        if (liabilityNet < 0m) liabilityNet = 0m;

                        decimal offset = Math.Min(assetNet, liabilityNet);
                        decimal difference = assetNet - liabilityNet;

                        // 坏账准备抵消（联动：仅被抵消的资产科目、债权方坏账准备，全额）
                        decimal badDebt = 0m;
                        if (hasBadDebtColumn && offset > 0m)
                            badDebtSums.TryGetValue(
                                creditor + "\u0001" + debtor + "\u0001" + years[y] + "\u0001" + assetAcc, out badDebt);

                        sequence++;
                        rowNumber++;
                        EliminationDiffRow line = new EliminationDiffRow();
                        line.Index = sequence;
                        line.Year = years[y];
                        line.Creditor = creditor;
                        line.Debtor = debtor;
                        line.AssetAccount = pairAcc.Asset;
                        line.LiabilityAccount = pairAcc.Liability;
                        line.RowNumber = rowNumber;
                        line.GroupIndex = groupIndex;
                        line.BadDebt = badDebt;
                        line.Asset = MakeCell(assetNet,
                            NetFormula(options, letters, years[y], assetAcc, liabAcc, creditor, debtor, true));
                        line.Liability = MakeCell(liabilityNet,
                            NetFormula(options, letters, years[y], assetAcc, liabAcc, creditor, debtor, false));
                        line.Offset = MakeCell(offset, OffsetFormula(rowNumber));
                        line.Difference = MakeCell(difference, DifferenceFormula(rowNumber));
                        line.OffsetReference = OffsetRef(SheetNameOf(options), rowNumber);
                        line.Note = NoteFor(assetNet, liabilityNet, difference);
                        diff.Rows.Add(line);
                    }
                }
                groupIndex++;
                // 每个公司对结束后空 N 行
                rowNumber += options.GroupGap > 0 ? options.GroupGap : 0;
            }
            diff.GroupCount = groupIndex;

            // ---- ③ 结果表：仅输出抵消金额不为 0 的科目对 ----
            BuildResult(result, diff, pairs, options);
            return;
        }

        // ---------------- 结果表 ----------------

        private static void BuildResult(EliminationMatrix result, EliminationDiffTable diff,
            IList<DiffAccountPair> pairs, MatrixOptions options)
        {
            result.Columns = EliminationMatrixBuilder.BuildColumns(diff.Years, options.StartColumn);
            // 结果表有第 1 行 SUBTOTAL + 两层表头
            int firstHeaderRow = options.StartRow + (options.TotalRow ? 1 : 0);
            int dataTop = firstHeaderRow + 2;
            int groupIndex = -1;
            int rowNumber = dataTop;
            int lastGroup = -1;

            // 分组：按 (公司对, 债权方, 债务方, 资产科目) 归并多年度（合并同类行），
            // 使坏账准备抵消两行能紧跟债权方行、插在债务方行之前。
            // 键含债权方/债务方，方向翻转的年度自然拆成不同组。
            Dictionary<string, List<EliminationDiffRow>> grouped = new Dictionary<string, List<EliminationDiffRow>>(StringComparer.Ordinal);
            List<string> groupOrder = new List<string>();

            foreach (EliminationDiffRow line in diff.Rows)
            {
                if (line.Offset == null || line.Offset.Value == 0m) continue;
                string key = line.GroupIndex + "\u0001" + line.Creditor + "\u0001" + line.Debtor + "\u0001" + line.AssetAccount;
                List<EliminationDiffRow> list;
                if (!grouped.TryGetValue(key, out list))
                {
                    list = new List<EliminationDiffRow>();
                    grouped[key] = list;
                    groupOrder.Add(key);
                }
                list.Add(line);
            }

            foreach (string key in groupOrder)
            {
                List<EliminationDiffRow> lines = grouped[key];
                EliminationDiffRow first = lines[0];
                if (first.GroupIndex != lastGroup)
                {
                    if (lastGroup >= 0) rowNumber += options.GroupGap > 0 ? options.GroupGap : 0;
                    lastGroup = first.GroupIndex;
                    groupIndex++;
                }

                // 债权方借资产科目（合并各年度）
                MatrixRow creditRow = new MatrixRow();
                creditRow.RowNumber = rowNumber++;
                creditRow.Entity = first.Creditor;
                creditRow.Customer = first.Debtor;
                creditRow.Account = first.AssetAccount;
                creditRow.Adjustment = "往来抵消";
                creditRow.TemplateKey = "ReceivableAsset";
                creditRow.Note = "";
                foreach (EliminationDiffRow line in lines)
                    creditRow.Debit[line.Year] = AsReference(line.Offset, line.OffsetReference);
                result.Rows.Add(creditRow);

                // 坏账准备抵消（借坏账准备 + 贷管理费用，全额、静态值、紧跟债权方行）
                bool hasBadDebt = false;
                foreach (EliminationDiffRow line in lines) if (line.BadDebt != 0m) { hasBadDebt = true; break; }
                if (hasBadDebt)
                {
                    MatrixRow badRow = new MatrixRow();
                    badRow.RowNumber = rowNumber++;
                    badRow.Entity = first.Creditor;
                    badRow.Customer = first.Debtor;
                    badRow.Account = BadDebtAccountName;
                    badRow.Adjustment = BadDebtAdjustment;
                    badRow.TemplateKey = "BadDebt";
                    badRow.Note = "";
                    foreach (EliminationDiffRow line in lines)
                        if (line.BadDebt != 0m) badRow.Debit[line.Year] = MakeCell(line.BadDebt, null);
                    result.Rows.Add(badRow);

                    MatrixRow mgmtRow = new MatrixRow();
                    mgmtRow.RowNumber = rowNumber++;
                    mgmtRow.Entity = first.Creditor;
                    mgmtRow.Customer = BadDebtMgmtCustomer;
                    mgmtRow.Account = BadDebtMgmtAccountName;
                    mgmtRow.Adjustment = BadDebtAdjustment;
                    mgmtRow.TemplateKey = "BadDebtMgmt";
                    mgmtRow.Note = "";
                    foreach (EliminationDiffRow line in lines)
                        if (line.BadDebt != 0m) mgmtRow.Credit[line.Year] = MakeCell(line.BadDebt, null);
                    result.Rows.Add(mgmtRow);
                }

                // 债务方贷负债科目（合并各年度）
                MatrixRow debtRow = new MatrixRow();
                debtRow.RowNumber = rowNumber++;
                debtRow.Entity = first.Debtor;
                debtRow.Customer = first.Creditor;
                debtRow.Account = first.LiabilityAccount;
                debtRow.Adjustment = "往来抵消";
                debtRow.TemplateKey = "ReceivableLiability";
                debtRow.Note = "";
                foreach (EliminationDiffRow line in lines)
                    debtRow.Credit[line.Year] = AsReference(line.Offset, line.OffsetReference);
                result.Rows.Add(debtRow);
            }

            // 结果表每个科目对 2 行（债权方借 / 债务方贷，多年度金额同列分栏），坏账准备抵消额外两行；
            // 主体对之间的空行已在 rowNumber 里加过（GroupStride 仅供收入成本场景的恒等式校验使用）
            result.TemplateCount = 2;
            result.GroupCount = groupIndex + 1;
            result.GroupGap = options.GroupGap;
            result.GroupStride = 2;
        }

        // ---------------- 取数 ----------------

        /// <summary>对方单位笔误别名 → 规范名（数据质量兜底，可扩展）。</summary>
        private static readonly Dictionary<string, string> PartyAliases =
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "芜湖市繁昌区繁花乡韵酒店管理有限公司", "芜湖市繁花乡韵酒店管理有限公司" },
                { "芜湖市繁昌区繁花食锦供应链北站分公司", "芜湖市繁昌区繁花食锦供应链管理有限公司北站分公司" },
                { "芜湖市繁昌区繁花食锦供应链管理有限公司繁昌北站分公司", "芜湖市繁昌区繁花食锦供应链管理有限公司北站分公司" },
                { "芜湖市繁昌区繁花食锦供应链管理有限公司繁昌慢谷分公司", "芜湖市繁昌区繁花食锦供应链管理有限公司慢谷分公司" },
            };

        /// <summary>把核算主体/对方单位的笔误别名归一化到规范名。</summary>
        private static string CanonicalName(string name)
        {
            if (name == null) return "";
            name = name.Trim();
            string canon;
            return PartyAliases.TryGetValue(name, out canon) ? canon : name;
        }

        /// <summary>规范名 → 其全部别名变体（不含规范名自身，用于公式的 OR 匹配）。</summary>
        private static List<string> AliasVariants(string canonicalName)
        {
            List<string> list = null;
            foreach (KeyValuePair<string, string> kv in PartyAliases)
                if (kv.Value == canonicalName)
                { if (list == null) list = new List<string>(); list.Add(kv.Key); }
            return list;
        }

        /// <summary>单向合计：主体=from、对方=to、指定年度、指定科目的审定数之和（精确匹配）。</summary>
        private static decimal DirectionalSum(Dictionary<string, decimal> sums,
            string from, string to, string year, string account)
        {
            if (account == null || account.Length == 0) return 0m;
            decimal total = 0m;
            foreach (KeyValuePair<string, decimal> entry in sums)
            {
                string[] parts = entry.Key.Split('\u0001');
                if (parts.Length < 4) continue;
                if (parts[0] == from && parts[1] == to && parts[2] == year && parts[3] == account)
                    total += entry.Value;
            }
            return total;
        }

        /// <summary>确定源表实际使用的科目名：优先输出名，其次别名；都不存在返回空串（金额为 0）。</summary>
        private static string SourceAccount(HashSet<string> accountValues, string name, string alias)
        {
            if (accountValues.Contains(name)) return name;
            if (alias != null && alias.Length > 0 && accountValues.Contains(alias)) return alias;
            return "";
        }

        // ---------------- 备注 ----------------

        /// <summary>
        /// 备注四分支。差额统一取「资产端账面净额 − 负债端账面净额」的绝对值。
        /// 建议按方向区分：资产方大=对方账上少记 → 核查对方；负债方大=我方账上少记 → 核查本单位。
        /// 依据合并实务：内部往来差异主因为未达账项（单据在途、入账时间差）、漏记/暂估、
        /// 串户/科目错挂、重分类错误。
        /// </summary>
        private static string NoteFor(decimal asset, decimal liability, decimal difference)
        {
            if (asset == 0m && liability == 0m) return "未查询到相关账户，余额为0";
            if (asset == liability) return "以小抵大：双方余额相等，全额抵消";
            string gap = Text(difference);   // |资产 − 负债|
            if (asset > liability)
                return "资产方大于负债方，差额" + gap + "为对方应收超额，建议核查对方未达账项（单据在途、入账时间差）、对方漏记入账及预收账款未重分类";
            return "负债方大于资产方，差额" + gap + "为对方应付超额，建议核查本单位未达账项（单据在途、入账时间差）、暂估未入账及预付账款未重分类";
        }

        private static string Text(decimal value)
        { return Math.Abs(value).ToString("#,##0.00", CultureInfo.InvariantCulture); }

        // ---------------- 单元格与公式 ----------------

        /// <summary>沿用金额值、但把公式换成跨表引用（引用缺失时退回原金额）。</summary>
        private static MatrixAmount AsReference(MatrixAmount source, string reference)
        {
            if (source == null) return null;
            if (string.IsNullOrEmpty(reference)) return source;
            MatrixAmount copy = new MatrixAmount();
            copy.Value = source.Value;
            copy.Formula = reference;
            return copy;
        }

        private static MatrixAmount MakeCell(decimal value, string formula)
        {
            MatrixAmount amount = new MatrixAmount();
            amount.Value = Math.Round(value, 2, MidpointRounding.AwayFromZero);
            amount.Formula = formula;
            return amount;
        }

        /// <summary>本行引用：抵消 = MIN(资产,负债)。差异表资产在 G 列、负债在 H 列（1 起：7、8）。</summary>
        private static string OffsetFormula(int rowNumber)
        { return "=MIN(G" + rowNumber + ",H" + rowNumber + ")"; }

        /// <summary>差异表表名（未指定时用默认名）。</summary>
        private static string SheetNameOf(MatrixOptions options)
        {
            string name = options == null ? "" : (options.DiffSheetName ?? "").Trim();
            return name.Length > 0 ? name : EliminationGateway.DiffDefaultSheetName;
        }

        /// <summary>
        /// 结果表金额用的跨表引用：='差异表'!I&lt;行号&gt;。
        /// 与 Python 参考实现一致（过程表 → 分录金额链接），避免结果表重算 SUMPRODUCT。
        /// 表名含中文/空格故加单引号；表名内的单引号按 Excel 规则双写。
        /// </summary>
        private static string OffsetRef(string sheetName, int rowNumber)
        {
            if (string.IsNullOrEmpty(sheetName)) return null;
            return "='" + sheetName.Replace("'", "''") + "'!I" + rowNumber;
        }

        /// <summary>差额 = 资产端账面净额 − 负债端账面净额（G 列资产、H 列负债）。</summary>
        private static string DifferenceFormula(int rowNumber)
        { return "=G" + rowNumber + "-H" + rowNumber; }

        /// <summary>
        /// 净额公式：= 单向合计(from→to, 科目1) − 单向合计(from→to, 科目2)。
        /// 资产端：from=债权方、to=债务方、科目1=资产科目、科目2=负债科目；
        /// 负债端：from=债务方、to=债权方、科目1=负债科目、科目2=资产科目。
        /// 科目在源表不存在时该侧合计为 0；两侧都不存在返回 null（回退静态值）。
        /// </summary>
        private static string NetFormula(MatrixOptions options, Letters letters,
            string year, string assetAcc, string liabAcc, string creditor, string debtor, bool isAsset)
        {
            string from = isAsset ? creditor : debtor;
            string to = isAsset ? debtor : creditor;
            string acc1 = isAsset ? assetAcc : liabAcc;
            string acc2 = isAsset ? liabAcc : assetAcc;

            string t1 = acc1 == null || acc1.Length == 0 ? "0"
                : DirectionalSumTerm(options, letters, year, acc1, from, to);
            string t2 = acc2 == null || acc2.Length == 0 ? "0"
                : DirectionalSumTerm(options, letters, year, acc2, from, to);
            if (t1 == null || t2 == null) return null;
            if (t1 == "0" && t2 == "0") return null;   // 无往来的科目对回退静态 0
            return "=" + t1 + "-" + t2;
        }

        /// <summary>
        /// 单向合计的活公式（不带前导 =）：
        /// SUMPRODUCT((主体=from)*(对方=to 及其别名)*(科目=account)*(LEFT(截止日,4)=year)*审定数)。
        /// 对方单位用「规范名 + 别名变体」的 OR 匹配，精确到公司名、不再前缀匹配（避免食锦/食锦北站被误并）。
        /// 条件数组与金额因子同区间（维度一致），避免 #N/A（实机验收教训）。
        /// </summary>
        private static string DirectionalSumTerm(MatrixOptions options, Letters letters,
            string year, string account, string fromName, string toName)
        {
            if (account == null || account.Length == 0) return null;
            FormulaKit kit = new FormulaKit(options);
            string entity = kit.ColData(letters.Entity);
            string party = kit.ColData(letters.Counterparty);
            string acc = kit.ColData(letters.Account);
            string deadline = kit.ColData(letters.Deadline);
            string amount = kit.ColData(letters.Amount);
            if (entity.Length == 0 || party.Length == 0 || acc.Length == 0
                || deadline.Length == 0 || amount.Length == 0) return null;

            StringBuilder sb = new StringBuilder("SUMPRODUCT(");
            sb.Append("(").Append(entity).Append("=\"").Append(fromName.Replace("\"", "\"\"")).Append("\")");
            sb.Append("*").Append(PartyMatchClause(party, toName));
            sb.Append("*(").Append(acc).Append("=\"").Append(account.Replace("\"", "\"\"")).Append("\")");
            sb.Append("*(LEFT(").Append(deadline).Append(",4)=\"").Append(year).Append("\")");
            sb.Append("*").Append(amount);
            sb.Append(")");
            return sb.ToString();
        }

        /// <summary>对方单位列的匹配子句：规范名 + 别名变体的 OR，形如 ((列="规范名")+(列="别名"))。</summary>
        private static string PartyMatchClause(string partyCol, string canonicalName)
        {
            List<string> variants = AliasVariants(canonicalName);
            StringBuilder sb = new StringBuilder("(");
            sb.Append("(").Append(partyCol).Append("=\"").Append(canonicalName.Replace("\"", "\"\"")).Append("\")");
            if (variants != null)
                foreach (string v in variants)
                    sb.Append("+(").Append(partyCol).Append("=\"").Append(v.Replace("\"", "\"\"")).Append("\")");
            sb.Append(")");
            return sb.ToString();
        }

        // ---------------- 工具 ----------------

        internal static string Text(EliminationRow row, EliminationRole role)
        {
            string value;
            return row.Values.TryGetValue(role, out value) ? (value ?? "").Trim() : "";
        }

        internal static int CompareYearDesc(string a, string b)
        {
            int ia, ib;
            bool oa = int.TryParse(a, NumberStyles.Integer, CultureInfo.InvariantCulture, out ia);
            bool ob = int.TryParse(b, NumberStyles.Integer, CultureInfo.InvariantCulture, out ib);
            if (oa && ob) return ib.CompareTo(ia);
            if (oa) return -1;
            if (ob) return 1;
            return string.Compare(b, a, StringComparison.Ordinal);
        }
    }
}
