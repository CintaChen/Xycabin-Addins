using System;
using System.Collections.Generic;
using System.Globalization;

namespace Excel_Iterate
{
    public static partial class EliminationEngine
    {
        /// <summary>收入成本抵消：逐行出分录（借营业收入 R；贷营业成本 R−U；贷存货 U）。</summary>
        public static EliminationResult BuildRevenueCost(IList<EliminationRow> rows)
        {
            EliminationResult result = new EliminationResult();
            if (rows == null || rows.Count == 0) return result;
            foreach (EliminationRow row in rows)
            {
                string unit = Text(row, EliminationRole.Unit);
                decimal revenue, cost, ratio;
                if (!TryParseAmount(Text(row, EliminationRole.Revenue), out revenue)
                    || !TryParseAmount(Text(row, EliminationRole.Cost), out cost)
                    || !TryParseRatio(Text(row, EliminationRole.UnsoldRatio), out ratio))
                {
                    result.SkippedRows++;
                    result.Warnings.Add("第 " + row.SourceRow + " 行收入、成本或未对外销售比例无法解析，已跳过。");
                    continue;
                }
                if (revenue <= 0m)
                {
                    result.SkippedRows++;
                    result.Warnings.Add("第 " + row.SourceRow + " 行收入为 " + Money(revenue) + "，无抵消意义，已跳过。");
                    continue;
                }
                decimal unrealized = Money2((revenue - cost) * ratio);
                decimal costBack = Money2(revenue - unrealized);
                string group = unit.Length > 0 ? unit : "第" + row.SourceRow + "行";
                EliminationEntry entry = new EliminationEntry();
                entry.Group = group;
                entry.Summary = "收入成本抵消：" + group;
                entry.Lines.Add(new EliminationLine
                {
                    IsDebit = true,
                    Account = AccRevenue,
                    Amount = Money2(revenue),
                    Note = "抵消内部交易收入 R=" + Money(revenue)
                });
                if (costBack > 0m)
                    entry.Lines.Add(new EliminationLine
                    {
                        IsDebit = false,
                        Account = AccCost,
                        Amount = costBack,
                        Note = "R−U=" + Money(revenue) + "−" + Money(unrealized) + "=" + Money(costBack)
                    });
                if (unrealized > 0m)
                    entry.Lines.Add(new EliminationLine
                    {
                        IsDebit = false,
                        Account = AccInventory,
                        Amount = unrealized,
                        Note = "U=(R−C)×比例=(" + Money(revenue) + "−" + Money(cost) + ")×" + RatioText(ratio) + "=" + Money(unrealized)
                    });
                result.Entries.Add(entry);
            }
            Renumber(result);
            return result;
        }
    }
}
