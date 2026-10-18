using System;
using System.Collections.Generic;
using System.Threading;
using WordUpdater;

// =====================================================================
// 「链接到EXCEL」映射规则离线自检（不需要 WPS）
//
// 验证两件事：
//   1. 源引用 → 一维单元格序列（含 Ctrl 多区域拆分）
//   2. 选中区域 → 目标矩形（含规则矩形 / 越界 / 跨表 判定）与行优先展开
//
// 假对象模型三要点（本仓惯例）：
//   · 假类全部 public —— dynamic 绑定受可访问性限制
//   · C# 索引器默认发出方法名 Item，同类不能既有 this[int] 又有 Item(int)
//   · 断言写「正确行为应该是什么」，而不是复述当前实现
// =====================================================================

public class FakeRange
{
    public int Start;
    public int End;
    public FakeTables Tables = new FakeTables(new List<FakeTableRef>());

    public FakeRange(int s, int e)
    {
        Start = s;
        End = e;
    }
}

public class FakeTableRef
{
    public FakeRange Range;

    public FakeTableRef(int start)
    {
        Range = new FakeRange(start, start + 100);
    }
}

public class FakeTables
{
    private List<FakeTableRef> _list;

    public FakeTables(List<FakeTableRef> list) { _list = list; }
    public int Count { get { return _list.Count; } }
    public FakeTableRef this[int i] { get { return _list[i - 1]; } }
}

public class FakeCell
{
    public int RowIndex;
    public int ColumnIndex;
    public FakeRange Range;
}

public class FakeCellCollection
{
    private List<FakeCell> _list;

    public FakeCellCollection(List<FakeCell> list) { _list = list; }
    public int Count { get { return _list.Count; } }
    public FakeCell this[int i] { get { return _list[i - 1]; } }
}

public class FakeCount
{
    private int _n;
    public FakeCount(int n) { _n = n; }
    public int Count { get { return _n; } }
}

public class FakeTable
{
    public FakeCount Rows;
    public FakeCount Columns;
    public FakeRange Range;
}

public class FakeSelection
{
    public FakeRange Range;
    public FakeCellCollection Cells;
}

internal static class LinkMapProbe
{
    private static int _pass;
    private static int _fail;

    private static void Check(string name, bool ok, string detail)
    {
        if (ok)
        {
            _pass++;
            Console.WriteLine("[PASS] " + name);
        }
        else
        {
            _fail++;
            Console.WriteLine("[FAIL] " + name + "   " + detail);
        }
    }

    private static void CheckRefs(string name, string fieldCode, string expected)
    {
        List<string> refs;
        string err;
        bool ok = LinkPasteEngine.TryListSourceRefs(fieldCode, out refs, out err);
        if (!ok)
        {
            Check(name, false, "解析失败：" + err);
            return;
        }
        string actual = string.Join(",", refs.ToArray());
        Check(name, actual == expected, "期望 [" + expected + "] 实际 [" + actual + "]");
    }

    private static void CheckRefsFail(string name, string fieldCode)
    {
        List<string> refs;
        string err;
        bool ok = LinkPasteEngine.TryListSourceRefs(fieldCode, out refs, out err);
        Check(name, !ok, "本应拒绝，却解析出了 " + refs.Count + " 条");
    }

    // ---------- 假对象构造 ----------

    private static FakeTable MakeTable(int rows, int cols)
    {
        FakeTable t = new FakeTable();
        t.Rows = new FakeCount(rows);
        t.Columns = new FakeCount(cols);
        t.Range = new FakeRange(0, 1000);
        return t;
    }

    private static FakeSelection MakeSel(int[] rowIdx, int[] colIdx, int[] parentStart,
                                         int selStart, int selEnd)
    {
        List<FakeCell> cs = new List<FakeCell>();
        for (int i = 0; i < rowIdx.Length; i++)
        {
            FakeCell c = new FakeCell();
            c.RowIndex = rowIdx[i];
            c.ColumnIndex = colIdx[i];
            int ps = (parentStart != null && i < parentStart.Length) ? parentStart[i] : 0;
            c.Range = new FakeRange(10 + i * 2, 11 + i * 2);
            List<FakeTableRef> tl = new List<FakeTableRef>();
            tl.Add(new FakeTableRef(ps));
            c.Range.Tables = new FakeTables(tl);
            cs.Add(c);
        }
        FakeSelection s = new FakeSelection();
        s.Cells = new FakeCellCollection(cs);
        s.Range = new FakeRange(selStart, selEnd);
        return s;
    }

    private static string RectText(LinkTargetRect r)
    {
        if (r == null) return "(null)";
        return r.Row1 + "," + r.Col1 + "," + r.Row2 + "," + r.Col2 + " (" + r.Rows + "x" + r.Cols + ")";
    }

    private static void CheckRect(string name, FakeSelection sel, FakeTable tbl,
                                  TargetRectStatus expectStatus, string expectRect)
    {
        LinkTargetRect rect;
        string detail;
        TargetRectStatus st = LinkPasteGeometry.ResolveTargetRect(sel, tbl, out rect, out detail);
        if (st != expectStatus)
        {
            Check(name, false, "期望状态 " + expectStatus + " 实际 " + st + "：" + detail);
            return;
        }
        if (expectRect == null)
        {
            Check(name, true, "");
            return;
        }
        string actual = RectText(rect);
        Check(name, actual == expectRect, "期望 [" + expectRect + "] 实际 [" + actual + "]");
    }

    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.WriteLine("== LinkMapProbe：链接映射规则离线自检 ==");
        Console.WriteLine();

        // ============ 1. 源引用 → 单元格序列 ============
        Console.WriteLine("--- 源引用展开（行优先）---");

        CheckRefs("连续区域 R1C1:R2C2 展成 4 格（先行后列）",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!R1C1:R2C2\" \\t \\a",
            "Sheet1!R1C1,Sheet1!R1C2,Sheet1!R2C1,Sheet1!R2C2");

        CheckRefs("连续区域 A1:B2 展成 4 格（先行后列）",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!A1:B2\" \\a",
            "Sheet1!A1,Sheet1!B1,Sheet1!A2,Sheet1!B2");

        CheckRefs("单行区域 A1:C1 展成 3 格",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!A1:C1\" \\a",
            "Sheet1!A1,Sheet1!B1,Sheet1!C1");

        CheckRefs("Ctrl 多区域 R1C1,R3C1（第二片省略工作表名，应继承）",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!R1C1,R3C1\" \\a",
            "Sheet1!R1C1,Sheet1!R3C1");

        CheckRefs("Ctrl 多区域 每片都带工作表名",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!R1C1,Sheet1!R3C1\" \\a",
            "Sheet1!R1C1,Sheet1!R3C1");

        CheckRefs("Ctrl 多区域 A1,B2 视为两个单格（不是 A1:B2 区域）",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!A1,B2\" \\a",
            "Sheet1!A1,Sheet1!B2");

        CheckRefs("Ctrl 多区域 H1,H3 映射为两个单格",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!H1,H3\" \\a",
            "Sheet1!H1,Sheet1!H3");

        CheckRefs("分号分隔同样视为多区域",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!A1;C1\" \\a",
            "Sheet1!A1,Sheet1!C1");

        CheckRefs("单个单元格 R11C12（L11 场景）",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!R11C12\" \\a",
            "Sheet1!R11C12");

        CheckRefs("DDE 形态无工作表前缀",
            "DDE Excel \"[C:\\a.xlsx]Sheet1\" \"R1C1:R1C2\"",
            "R1C1,R1C2");

        CheckRefsFail("无法识别的引用应被拒绝",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"ZZZ\" \\a");

        CheckRefsFail("空代码应被拒绝", "");

        // ============ 2. 目标矩形解析 ============
        Console.WriteLine();
        Console.WriteLine("--- 目标矩形解析 ---");

        FakeTable t3 = MakeTable(3, 3);

        // 只选 (3,3) 一格
        CheckRect("3x3 表只选 (3,3) 一格 → 1x1",
            MakeSel(new int[] { 3 }, new int[] { 3 }, null, 10, 12), t3,
            TargetRectStatus.Ok, "3,3,3,3 (1x1)");

        // 选 (1,1)-(2,2)
        CheckRect("选中 (1,1)-(2,2) → 2x2",
            MakeSel(new int[] { 1, 1, 2, 2 }, new int[] { 1, 2, 1, 2 }, null, 10, 20), t3,
            TargetRectStatus.Ok, "1,1,2,2 (2x2)");

        // 选中第 1 列第 1~2 行（用户场景：H1/H3 → 两行一列）
        CheckRect("选中第 1 列的第 1~2 行 → 2x1",
            MakeSel(new int[] { 1, 2 }, new int[] { 1, 1 }, null, 10, 20), t3,
            TargetRectStatus.Ok, "1,1,2,1 (2x1)");

        // 选中第 1 行第 1~2 列（用户场景：A1/C1 → 一行两列）
        CheckRect("选中第 1 行的第 1~2 列 → 1x2",
            MakeSel(new int[] { 1, 1 }, new int[] { 1, 2 }, null, 10, 20), t3,
            TargetRectStatus.Ok, "1,1,1,2 (1x2)");

        // 全表
        int[] allR = new int[9];
        int[] allC = new int[9];
        int k = 0;
        for (int r = 1; r <= 3; r++)
            for (int c = 1; c <= 3; c++) { allR[k] = r; allC[k] = c; k++; }
        CheckRect("选中整张 3x3 表 → 3x3",
            MakeSel(allR, allC, null, 0, 1000), t3,
            TargetRectStatus.Ok, "1,1,3,3 (3x3)");

        // L 形（不是矩形）
        CheckRect("L 形选区应判定不是规则矩形",
            MakeSel(new int[] { 1, 1, 2, 3 }, new int[] { 1, 2, 1, 1 }, null, 10, 20), t3,
            TargetRectStatus.NotRect, null);

        // 越出表格
        CheckRect("选区跨出表格边界应判定越界",
            MakeSel(new int[] { 3 }, new int[] { 3 }, null, 10, 5000), t3,
            TargetRectStatus.OutOfTable, null);

        // 行号超出表格行数
        CheckRect("单元格行号超出表格行数应判定越界",
            MakeSel(new int[] { 9 }, new int[] { 1 }, null, 10, 20), t3,
            TargetRectStatus.OutOfTable, null);

        // 跨两张表
        CheckRect("选区横跨两张表格应判定跨表",
            MakeSel(new int[] { 1, 1 }, new int[] { 1, 2 }, new int[] { 0, 500 }, 10, 20), t3,
            TargetRectStatus.MultiTable, null);

        // 没有单元格
        CheckRect("选区没有单元格应判定 NoCells",
            MakeSel(new int[0], new int[0], null, 10, 20), t3,
            TargetRectStatus.NoCells, null);

        // ============ 3. 目标行优先展开顺序 ============
        Console.WriteLine();
        Console.WriteLine("--- 目标行优先展开顺序 ---");

        LinkTargetRect rect22 = new LinkTargetRect();
        rect22.Row1 = 1; rect22.Col1 = 1; rect22.Row2 = 2; rect22.Col2 = 2;
        List<LinkCellPos> flat = LinkPasteGeometry.FlattenTarget(rect22);
        string order = "";
        for (int i = 0; i < flat.Count; i++)
        {
            if (i > 0) order += ",";
            order += "(" + flat[i].Row + "," + flat[i].Col + ")";
        }
        Check("2x2 展开顺序为 (1,1),(1,2),(2,1),(2,2)",
            order == "(1,1),(1,2),(2,1),(2,2)", "实际 " + order);

        // 用户场景：源 H1,H3 → 目标第 1 列第 1~2 行
        LinkTargetRect col2 = new LinkTargetRect();
        col2.Row1 = 1; col2.Col1 = 1; col2.Row2 = 2; col2.Col2 = 1;
        List<LinkCellPos> f2 = LinkPasteGeometry.FlattenTarget(col2);
        Check("H1,H3 → 第1列1~2行：目标序列长度=2 且为 (1,1),(2,1)",
            f2.Count == 2 && f2[0].Row == 1 && f2[0].Col == 1 && f2[1].Row == 2 && f2[1].Col == 1,
            "实际 " + f2.Count + " 格");

        // 用户场景：源 A1,C1 → 目标第 1 行第 1~2 列
        LinkTargetRect row2 = new LinkTargetRect();
        row2.Row1 = 1; row2.Col1 = 1; row2.Row2 = 1; row2.Col2 = 2;
        List<LinkCellPos> f3 = LinkPasteGeometry.FlattenTarget(row2);
        Check("A1,C1 → 第1行1~2列：目标序列长度=2 且为 (1,1),(1,2)",
            f3.Count == 2 && f3[0].Row == 1 && f3[0].Col == 1 && f3[1].Row == 1 && f3[1].Col == 2,
            "实际 " + f3.Count + " 格");

        // ============ 4. 数量匹配（用户定义的严格规则）============
        Console.WriteLine();
        Console.WriteLine("--- 数量严格匹配 ---");

        List<string> refsA;
        string errA;
        LinkPasteEngine.TryListSourceRefs(
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!A1:B2\" \\a", out refsA, out errA);
        LinkTargetRect one = new LinkTargetRect();
        one.Row1 = 3; one.Col1 = 3; one.Row2 = 3; one.Col2 = 3;
        Check("源 2x2(4 格) 对目标 (3,3) 单格(1 格) → 数量不等，应拒绝",
            refsA.Count == 4 && Flatten(one) == 1 && refsA.Count != Flatten(one),
            "源=" + refsA.Count + " 目标=" + Flatten(one));

        List<string> refsB;
        string errB;
        LinkPasteEngine.TryListSourceRefs(
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!R11C12\" \\a", out refsB, out errB);
        Check("源单格 L11(1 格) 对目标 (3,3) 单格(1 格) → 数量相等，应放行",
            refsB.Count == 1 && Flatten(one) == 1 && refsB.Count == Flatten(one),
            "源=" + refsB.Count + " 目标=" + Flatten(one));

        Console.WriteLine();
        Console.WriteLine("结果：PASS=" + _pass + "  FAIL=" + _fail);
        Console.Out.Flush();
        return _fail == 0 ? 0 : 1;
    }

    private static int Flatten(LinkTargetRect r)
    {
        return LinkPasteGeometry.FlattenTarget(r).Count;
    }
}
