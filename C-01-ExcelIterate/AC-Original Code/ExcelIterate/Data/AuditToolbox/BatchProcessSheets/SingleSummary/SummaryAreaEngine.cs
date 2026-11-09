using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    internal sealed class SummaryLeadRow
    {
        internal int Row, CodeColumn, NameColumn;
    }
    internal sealed class SummaryAreaCell
    {
        internal int Row, Column;
        internal string Text, R1C1;
        internal bool Formula;
    }
    internal static class SummaryAreaEngine
    {
        internal const int CellLimit = 100000;
        // Pair actual row identities rather than zipping flattened Areas: a
        // missing cell in either column must never shift subsequent companies.
        internal static List<SummaryLeadRow> LeadRows(IList<SummaryAreaPlan> areas)
        {
            if (areas.Count == 0) throw new InvalidOperationException("请选择 Lead 公司字段。");
            var columns = new SortedSet<int>(); var rows = new Dictionary<int, HashSet<int>>(); var order = new List<int>(); long count = 0;
            foreach (var area in areas)
            {
                if (area.Sheet != areas[0].Sheet) throw new InvalidOperationException("Lead 选区必须来自同一工作表。");
                if (area.Rows < 1 || area.Columns < 1 || area.Columns > 2 || area.Row < 1 || area.Column < 1 || (long)area.Row + area.Rows - 1 > 1048576 || (long)area.Column + area.Columns - 1 > 16384)
                    throw new InvalidOperationException("Lead 每个 Area 只能包含一列或两列的有效数据区域。");
                count += (long)area.Rows * area.Columns;
                if (count > CellLimit) throw new InvalidOperationException("Lead 选区最多十万单元格，请选择实际数据行。");
                for (int c = area.Column; c < area.Column + area.Columns; c++) columns.Add(c);
                if (columns.Count > 2) throw new InvalidOperationException("Lead 只能选择两列：左列编码简称，右列公司全称。");
                for (int r = area.Row; r < area.Row + area.Rows; r++)
                {
                    HashSet<int> selected;
                    if (!rows.TryGetValue(r, out selected)) { selected = new HashSet<int>(); rows.Add(r, selected); order.Add(r); }
                    for (int c = area.Column; c < area.Column + area.Columns; c++)
                        if (!selected.Add(c)) throw new InvalidOperationException("Lead 第 " + r + " 行有重复或重叠选区。");
                }
            }
            if (columns.Count != 2) throw new InvalidOperationException("Lead 需要两列：可选连续两列，也可 Ctrl 选择不相邻的两列。");
            var result = new List<SummaryLeadRow>();
            foreach (int row in order)
            {
                if (rows[row].Count != 2) throw new InvalidOperationException("Lead 第 " + row + " 行未同时选中编码和全称，请按同一行补齐两列。");
                result.Add(new SummaryLeadRow { Row = row, CodeColumn = columns.Min, NameColumn = columns.Max });
            }
            return result;
        }
        internal static bool Intersects(SummaryAreaPlan a, SummaryAreaPlan b)
        { return a.Row <= (long)b.Row + b.Rows - 1 && b.Row <= (long)a.Row + a.Rows - 1 && a.Column <= (long)b.Column + b.Columns - 1 && b.Column <= (long)a.Column + a.Columns - 1; }
        internal static List<SummaryAreaPlan> Targets(IList<SummaryAreaPlan> areas, int pasteRow, int pasteColumn, int companies)
        {
            if (areas.Count == 0 || companies < 1) throw new InvalidOperationException("请选择模板区域和至少一家公司。");
            int minRow = areas.Min(a => a.Row), minColumn = areas.Min(a => a.Column);
            int width = areas.Max(a => checked(a.Column + a.Columns - 1)) - minColumn + 1;
            long count = 0;
            for (int i = 0; i < areas.Count; i++)
            {
                var a = areas[i];
                if (a.Rows < 1 || a.Columns < 1 || a.Row < 1 || a.Column < 1 || (long)a.Row + a.Rows - 1 > 1048576 || (long)a.Column + a.Columns - 1 > 16384)
                    throw new InvalidOperationException("模板区域坐标无效。");
                if (a.Sheet != areas[0].Sheet) throw new InvalidOperationException("模板不能跨工作表。");
                count += (long)a.Rows * a.Columns;
                for (int j = 0; j < i; j++) if (Intersects(a, areas[j])) throw new InvalidOperationException("模板 Area 相互重叠。");
            }
            if (count * companies > CellLimit) throw new InvalidOperationException("MVP 每批目标最多十万单元格，请缩小区域或公司数量。");
            var result = new List<SummaryAreaPlan>();
            for (int company = 0; company < companies; company++) foreach (var area in areas)
            {
                var target = area.Target(minRow, minColumn, pasteRow, pasteColumn, width, company);
                if (areas.Any(source => Intersects(source, target))) throw new InvalidOperationException("模板与目标重叠，原始模板必须保持不变。");
                result.Add(target);
            }
            return result;
        }
        internal static void NativeFormula(string formula)
        {
            if (string.IsNullOrEmpty(formula) || formula[0] != '=') throw new InvalidOperationException("公式类型无法确认。");
            string mask = Regex.Replace(formula, "\"(?:[^\"]|\"\")*\"", "\"\"");
            if (mask.IndexOfAny(new[] { '[', ']', '!', '#', '@', '{', '}', '\'', ';' }) >= 0)
                throw new InvalidOperationException("已有外链、跨表、数组、溢出或特殊公式不在 MVP 原生复制白名单内。");
            mask = Regex.Replace(mask, @"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*", "0");
            var supported = new HashSet<string>(new[] { "ROUND", "SUM", "IF", "IFERROR", "ABS", "MIN", "MAX", "AVERAGE", "COUNT", "COUNTA", "AND", "OR", "NOT", "TRUE", "FALSE" });
            foreach (Match token in Regex.Matches(mask, @"[A-Za-z_][A-Za-z0-9_.]*"))
                if (!supported.Contains(token.Value.ToUpperInvariant())) throw new InvalidOperationException("已有公式含未确认名称或函数：" + token.Value);
        }
        internal static string Address(SummaryAreaPlan area)
        { return CellAddress.Name(area.Row, area.Column) + ":" + CellAddress.Name(area.Row + area.Rows - 1, area.Column + area.Columns - 1); }
        internal static void ValidateShift(string formula, int rows, int columns)
        {
            NativeFormula(formula);
            foreach (var span in FormulaRefRewriter.Scan(formula))
            {
                CheckReference(span.Row1, span.Column1, span.Row1Absolute, span.Column1Absolute, rows, columns);
                if (span.IsRange) CheckReference(span.Row2, span.Column2, span.Row2Absolute, span.Column2Absolute, rows, columns);
            }
        }
        private static void CheckReference(int row, int column, bool rowAbsolute, bool columnAbsolute, int deltaRow, int deltaColumn)
        {
            long r = rowAbsolute ? row : (long)row + deltaRow, c = columnAbsolute ? column : (long)column + deltaColumn;
            if (r < 1 || r > 1048576 || c < 1 || c > 16384) throw new InvalidOperationException("已有公式复制后会产生越界引用。");
        }
    }
}
