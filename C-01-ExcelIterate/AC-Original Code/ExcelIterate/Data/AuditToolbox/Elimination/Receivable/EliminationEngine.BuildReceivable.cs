using System;
using System.Collections.Generic;
using System.Globalization;

namespace Excel_Iterate
{
    public static partial class EliminationEngine
    {
        /// <summary>
        /// 往来抵消：按单位聚合多行，逐单位出分录。
        /// includeBadDebt=true 时才生成坏账准备冲回分录（映射含坏账准备列）。
        /// </summary>
        public static EliminationResult BuildReceivable(IList<EliminationRow> rows, bool includeBadDebt)
        {
            EliminationResult result = new EliminationResult();
            if (rows == null || rows.Count == 0) return result;

            // 按单位聚合（单位名去首尾空格后精确匹配）
            List<string> order = new List<string>();
            Dictionary<string, UnitSum> sums = new Dictionary<string, UnitSum>();
            foreach (EliminationRow row in rows)
            {
                string unit = Text(row, EliminationRole.Unit);
                if (unit.Length == 0)
                {
                    result.SkippedRows++;
                    result.Warnings.Add("第 " + row.SourceRow + " 行缺少单位名称，已跳过。");
                    continue;
                }
                decimal receivable, payable, badDebt = 0m;
                if (!TryParseAmount(Text(row, EliminationRole.Receivable), out receivable)
                    || !TryParseAmount(Text(row, EliminationRole.Payable), out payable))
                {
                    result.SkippedRows++;
                    result.Warnings.Add("第 " + row.SourceRow + " 行（" + unit + "）应收或应付金额无法解析，已跳过。");
                    continue;
                }
                UnitSum sum;
                if (!sums.TryGetValue(unit, out sum))
                {
                    sum = new UnitSum();
                    sums[unit] = sum;
                    order.Add(unit);
                }
                sum.Receivable += receivable;
                sum.Payable += payable;
                if (includeBadDebt)
                {
                    decimal bad;
                    if (TryParseAmount(Text(row, EliminationRole.BadDebt), out bad)) sum.BadDebt += bad;
                }
                sum.Rows++;
            }

            foreach (string unit in order)
            {
                UnitSum sum = sums[unit];
                if (sum.Receivable <= 0m || sum.Payable <= 0m)
                {
                    result.SkippedRows += sum.Rows;
                    result.Warnings.Add("单位「" + unit + "」应收 " + Money(sum.Receivable)
                        + " / 应付 " + Money(sum.Payable) + "，无内部对冲额，已跳过。");
                    continue;
                }
                decimal offset = Math.Min(sum.Receivable, sum.Payable);
                EliminationEntry entry = new EliminationEntry();
                entry.Group = unit;
                entry.Summary = "往来抵消：" + unit;
                entry.Lines.Add(new EliminationLine
                {
                    IsDebit = true,
                    Account = AccPayable,
                    Amount = Money2(offset),
                    Note = "对冲额=min(应收" + Money(sum.Receivable) + ",应付" + Money(sum.Payable) + ")=" + Money(offset)
                });
                entry.Lines.Add(new EliminationLine
                {
                    IsDebit = false,
                    Account = AccReceivable,
                    Amount = Money2(offset),
                    Note = "对冲额=min(应收" + Money(sum.Receivable) + ",应付" + Money(sum.Payable) + ")=" + Money(offset)
                });
                result.Entries.Add(entry);

                // 坏账准备按对冲比例冲回：仅当对冲后仍有剩余应收（应收>应付）时
                if (includeBadDebt && sum.Receivable > sum.Payable && sum.BadDebt > 0m)
                {
                    decimal reversal = Money2(sum.BadDebt * offset / sum.Receivable);
                    if (reversal > 0m)
                    {
                        EliminationEntry back = new EliminationEntry();
                        back.Group = unit;
                        back.Summary = "坏账准备冲回：" + unit;
                        back.Lines.Add(new EliminationLine
                        {
                            IsDebit = true,
                            Account = AccBadDebt,
                            Amount = reversal,
                            Note = "坏账冲回=坏账准备" + Money(sum.BadDebt) + "×" + Money(offset) + "/" + Money(sum.Receivable) + "=" + Money(reversal)
                        });
                        back.Lines.Add(new EliminationLine
                        {
                            IsDebit = false,
                            Account = AccCreditLoss,
                            Amount = reversal,
                            Note = "坏账冲回=坏账准备" + Money(sum.BadDebt) + "×" + Money(offset) + "/" + Money(sum.Receivable) + "=" + Money(reversal)
                        });
                        result.Entries.Add(back);
                    }
                }
            }
            Renumber(result);
            return result;
        }
    }
}
