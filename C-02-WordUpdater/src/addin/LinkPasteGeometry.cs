using System;
using System.Collections.Generic;

namespace WordUpdater
{
    // =====================================================================
    // 「链接Excel」的几何与序列工具 —— 纯逻辑、无 UI，便于离线探针自检。
    //
    // 核心模型：把「源」和「目标」都展成一维序列，按位一一对应。
    //   · 源序列  ：子区域按出现顺序展开，每个区域内先行后列（行优先）
    //   · 目标序列：选中矩形先行后列（行优先）展开
    //   · 两序列长度必须严格相等才写入；不等即拒绝，不做任何扩展/兜底
    //
    // 这一条规则同时覆盖了：单格、连续区域、Ctrl 多区域三种源形态，
    // 且连续区域的展开结果与改造前的逐格偏移写法完全一致。
    // =====================================================================

    /// <summary>目标矩形的行列号（1 起始、含端点）。用行列号而非 Range 位置保存，
    /// 因为后续剪贴板探测会在文档末尾做「粘贴→读取→删除」，Range 会漂移而行列号不会。</summary>
    public sealed class LinkTargetRect
    {
        public int Row1;   // 左上角行
        public int Col1;   // 左上角列
        public int Row2;   // 右下角行
        public int Col2;   // 右下角列

        public int Rows { get { return Row2 - Row1 + 1; } }
        public int Cols { get { return Col2 - Col1 + 1; } }
        public int CellCount { get { return Rows * Cols; } }

        public override string ToString()
        {
            if (CellCount == 1)
                return "第 " + Row1 + " 行第 " + Col1 + " 列";
            return "第 " + Row1 + "~" + Row2 + " 行 × 第 " + Col1 + "~" + Col2 + " 列";
        }
    }

    /// <summary>目标侧的一个席位（绝对行列号）。</summary>
    public sealed class LinkCellPos
    {
        public int Row;
        public int Col;

        public LinkCellPos(int row, int col)
        {
            Row = row;
            Col = col;
        }
    }

    /// <summary>源侧的一个单元格（已行优先展开）。</summary>
    public sealed class LinkSourceCell
    {
        public string SheetPart = "";   // "Sheet1!"；DDE 形态为空串
        public int Row;
        public int Col;
        public bool A1Style;            // true=A1 记法，false=R1C1 记法

        public string Ref()
        {
            if (A1Style) return LinkPasteGeometry.ColToLetters(Col) + Row.ToString();
            return "R" + Row.ToString() + "C" + Col.ToString();
        }

        public override string ToString()
        {
            return SheetPart + Ref();
        }
    }

    /// <summary>解析「用户选中的单元格矩形」的结果。</summary>
    public enum TargetRectStatus
    {
        Ok,          // 规则矩形
        ReadError,   // COM 读取失败（接口缺失或疑似合并单元格）
        NoCells,     // 选区里一个单元格都没有
        OutOfTable,  // 选区跨出表格边界（含表格外正文）
        MultiTable,  // 选区横跨多张表格
        NotRect      // 不是规则矩形
    }

    public static class LinkPasteGeometry
    {
        /// <summary>
        /// 读取「用户选中的单元格矩形」。只返回行列号，不保存 Range 对象。
        /// 所有 COM 调用单点 try/catch：WPS 常在各处抛异常，且合并单元格会让行列号读不出来。
        /// </summary>
        public static TargetRectStatus ResolveTargetRect(object selection, object table,
                                                         out LinkTargetRect rect, out string detail)
        {
            rect = null;
            detail = null;
            if (selection == null || table == null)
            {
                detail = "无法获取选区或目标表格。";
                return TargetRectStatus.ReadError;
            }

            dynamic sel = selection;
            dynamic tbl = table;

            // ---------- 1. 表格边界 ----------
            int tblStart = -1, tblEnd = -1, tblRows = -1, tblCols = -1;
            try { tblRows = (int)tbl.Rows.Count; } catch { }
            try { tblCols = (int)tbl.Columns.Count; } catch { }
            try { tblStart = (int)tbl.Range.Start; } catch { }
            try { tblEnd = (int)tbl.Range.End; } catch { }

            // ---------- 2. 选区必须完全落在这张表格内 ----------
            int s = -1, e = -1;
            try
            {
                dynamic sr = sel.Range;
                s = (int)sr.Start;
                e = (int)sr.End;
            }
            catch { }
            if (tblStart >= 0 && tblEnd >= 0 && s >= 0 && e >= 0 && (s < tblStart || e > tblEnd))
            {
                detail = "选区超出了这张表格的范围（包含了表格外的正文）。";
                return TargetRectStatus.OutOfTable;
            }

            // ---------- 3. 取选区单元格集合：sel.Cells 优先，sel.Range.Cells 兜底 ----------
            dynamic cells = null;
            int n = -1;
            try { cells = sel.Cells; n = (int)cells.Count; } catch { n = -1; }
            if (n < 0)
            {
                try { cells = sel.Range.Cells; n = (int)cells.Count; } catch { n = -1; }
            }
            if (n < 0)
            {
                detail = "无法读取选区中的单元格（Selection.Cells 不可用）。";
                return TargetRectStatus.ReadError;
            }
            if (n == 0)
            {
                detail = "没有选中任何表格单元格。";
                return TargetRectStatus.NoCells;
            }

            // ---------- 4. 逐格求包围盒（去重 + 父表唯一性） ----------
            HashSet<int> seen = new HashSet<int>();
            HashSet<int> parents = new HashSet<int>();
            int minRow = int.MaxValue, maxRow = -1, minCol = int.MaxValue, maxCol = -1;
            int got = 0;

            for (int i = 1; i <= n; i++)
            {
                dynamic cell = null;
                try { cell = UpdateEngine.GetItem(cells, i); } catch { continue; }
                if (cell == null) continue;

                int r = -1, c = -1;
                try { r = (int)cell.RowIndex; } catch { }
                try { c = (int)cell.ColumnIndex; } catch { }
                if (r <= 0 || c <= 0)
                {
                    detail = "读取第 " + i + " 个单元格的行列号失败（表格可能含合并单元格）。";
                    return TargetRectStatus.ReadError;
                }
                if (tblRows > 0 && r > tblRows)
                {
                    detail = "选区中的单元格（第 " + r + " 行）超出目标表格范围（共 " + tblRows + " 行）。";
                    return TargetRectStatus.OutOfTable;
                }
                if (tblCols > 0 && c > tblCols)
                {
                    detail = "选区中的单元格（第 " + c + " 列）超出目标表格范围（共 " + tblCols + " 列）。";
                    return TargetRectStatus.OutOfTable;
                }

                int key = r * 1000000 + c;
                if (!seen.Add(key)) continue;   // 同一格被重复计入时只算一次
                got++;
                if (r < minRow) minRow = r;
                if (r > maxRow) maxRow = r;
                if (c < minCol) minCol = c;
                if (c > maxCol) maxCol = c;

                // 父表判别：取每格所属表格的起点，出现多个起点即跨表
                int p = -1;
                try
                {
                    dynamic pt = UpdateEngine.GetItem(cell.Range.Tables, 1);
                    p = (int)pt.Range.Start;
                }
                catch { }
                if (p >= 0) parents.Add(p);
            }

            if (got == 0)
            {
                detail = "没有读取到任何可用的单元格。";
                return TargetRectStatus.NoCells;
            }
            if (parents.Count > 1)
            {
                detail = "选区横跨了 " + parents.Count + " 张表格。";
                return TargetRectStatus.MultiTable;
            }

            // ---------- 5. 规则矩形校验：格数必须恰好等于行列跨度 ----------
            int spanRows = maxRow - minRow + 1;
            int spanCols = maxCol - minCol + 1;
            if (got != spanRows * spanCols)
            {
                detail = string.Format(
                    "选区不是规则的矩形：共 {0} 个单元格，但行列跨度为 {1} 行 × {2} 列（{3} 格）。",
                    got, spanRows, spanCols, spanRows * spanCols);
                return TargetRectStatus.NotRect;
            }

            LinkTargetRect res = new LinkTargetRect();
            res.Row1 = minRow;
            res.Col1 = minCol;
            res.Row2 = maxRow;
            res.Col2 = maxCol;
            rect = res;
            return TargetRectStatus.Ok;
        }

        /// <summary>目标矩形按行优先（行内先换列，再换行）展成一维序列。</summary>
        public static List<LinkCellPos> FlattenTarget(LinkTargetRect rect)
        {
            List<LinkCellPos> list = new List<LinkCellPos>();
            if (rect == null) return list;
            for (int r = rect.Row1; r <= rect.Row2; r++)
            {
                for (int c = rect.Col1; c <= rect.Col2; c++)
                    list.Add(new LinkCellPos(r, c));
            }
            return list;
        }

        /// <summary>列号 → 列字母（1 → A，27 → AA）。</summary>
        internal static string ColToLetters(int col)
        {
            string s = "";
            while (col > 0)
            {
                int m = (col - 1) % 26;
                s = ((char)('A' + m)).ToString() + s;
                col = (col - 1) / 26;
            }
            return s;
        }
    }
}
