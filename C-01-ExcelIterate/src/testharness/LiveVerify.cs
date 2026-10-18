using System;
using System.Collections.Generic;
using System.Globalization;
using Excel_Iterate;

// =====================================================================
// 实机验证（临时工具，不属于交付物，不入 project.json）
//
// 目的：把插件真实生成的公式交给 WPS 表格计算，验证：
//   ① 收入成本抵消：SUMIFS（整列）、成本还原引用、SUMPRODUCT 现算（区间）
//   ② 往来抵消：SUMPRODUCT 精确匹配 + 别名 OR（净额法）
// 这是前两次只在真机才暴露问题（#VALUE!、#N/A）的核心防线。
//
// 编译：csc /target:exe /r:System.dll /r:System.Core.dll /r:Microsoft.CSharp.dll
//       /r:build\Excel_Iterate.dll LiveVerify.cs
// =====================================================================
internal static class LiveVerify
{
    const string Path = @"E:\Xymircai\A.1 CodeProject\B-1-ProgramDevelop\B-3-Wps\B-99-实验文档\合并抵消\实机验证.xlsx";
    static int pass = 0, fail = 0;

    static void Check(string name, bool ok, string detail)
    {
        if (ok) { pass++; Console.WriteLine("  [PASS] " + name); }
        else { fail++; Console.WriteLine("  [FAIL] " + name + " — " + detail); }
    }

    static EliminationRow R(int row, params object[] kv)
    {
        EliminationRow r = new EliminationRow();
        r.SourceRow = row;
        for (int i = 0; i + 1 < kv.Length; i += 2)
            r.Values[(EliminationRole)kv[i]] = Convert.ToString(kv[i + 1], CultureInfo.InvariantCulture);
        return r;
    }

    static EliminationColumn C(int number, string letter, string header, EliminationRole role)
    { return new EliminationColumn { Number = number, Letter = letter, Header = header, Role = role }; }

    /// <summary>删掉同名旧表后新建（重复运行时避免重名报错）。</summary>
    static dynamic FreshSheet(dynamic book, string name)
    {
        dynamic app = book.Application;
        bool old = false;
        try { old = app.DisplayAlerts; app.DisplayAlerts = false; } catch { }
        try
        {
            int count = book.Worksheets.Count;
            for (int i = count; i >= 1; i--)
            {
                dynamic s = book.Worksheets[i];
                if (string.Equals(Convert.ToString(s.Name), name, StringComparison.Ordinal))
                { s.Delete(); }
            }
        }
        catch { }
        try { app.DisplayAlerts = old; } catch { }
        dynamic sheet = book.Worksheets.Add();
        sheet.Name = name;
        return sheet;
    }

    /// <summary>字符串按公式写入，数值按值写入（与插件写入逻辑一致）。</summary>
    static void Put(dynamic sheet, string addr, object value)
    {
        if (value == null) return;
        string formula = value as string;
        if (formula != null) sheet.Range[addr].Formula = formula;
        else sheet.Range[addr].Value2 = value;
    }

    static void Main()
    {
        Console.WriteLine("=== 实机验证：WPS 表格计算插件生成的公式 ===");
        Type t = Type.GetTypeFromProgID("KET.Application");
        if (t == null) { Console.WriteLine("未找到 WPS 表格 COM（KET.Application）"); return; }

        dynamic app = null;
        try
        {
            app = Activator.CreateInstance(t);
            app.Visible = true;
            app.DisplayAlerts = false;
            dynamic book = app.Workbooks.Open(Path);

            VerifyRevenueCost(app, book);
            VerifyReceivable(app, book);

            try { book.Close(true); } catch { }
            try { app.Quit(); } catch { }
        }
        catch (Exception ex)
        {
            Console.WriteLine("异常：" + ex.Message + "\r\n" + ex.StackTrace);
            try { if (app != null) app.Quit(); } catch { }
        }
        Console.WriteLine();
        Console.WriteLine("=== 结果：" + pass + " PASS / " + fail + " FAIL ===");
    }

    // ---------------- 收入成本抵消 ----------------
    static void VerifyRevenueCost(dynamic app, dynamic book)
    {
        Console.WriteLine();
        Console.WriteLine("--- 收入成本抵消（矩阵式底稿）---");
        List<EliminationColumn> cols = new List<EliminationColumn>();
        cols.Add(C(1, "A", "核算主体", EliminationRole.Entity));
        cols.Add(C(5, "E", "年份", EliminationRole.Year));
        cols.Add(C(10, "J", "客户名称", EliminationRole.Counterparty));
        cols.Add(C(14, "N", "是否为合并内关联方", EliminationRole.Condition));
        cols.Add(C(15, "O", "单体收入（未审）", EliminationRole.Revenue));
        cols.Add(C(16, "P", "单体成本（未审）", EliminationRole.Cost));
        cols.Add(C(22, "V", "未对外销售比例", EliminationRole.UnsoldRatio));
        cols.Add(C(23, "W", "未实现内部交易损益（存货）", EliminationRole.UnrealizedProfit));

        List<EliminationRow> rows = new List<EliminationRow>();
        rows.Add(R(3, EliminationRole.Entity, "宿州一号", EliminationRole.Year, "2024",
            EliminationRole.Counterparty, "宿州十一号", EliminationRole.Condition, "是",
            EliminationRole.Revenue, "2217860.59", EliminationRole.Cost, "1849290.40",
            EliminationRole.UnsoldRatio, "20%"));
        rows.Add(R(4, EliminationRole.Entity, "宿州一号", EliminationRole.Year, "2025",
            EliminationRole.Counterparty, "宿州十一号", EliminationRole.Condition, "是",
            EliminationRole.Revenue, "1998483.64", EliminationRole.Cost, "2258286.52",
            EliminationRole.UnsoldRatio, "20%", EliminationRole.UnrealizedProfit, "-46492.5225"));
        rows.Add(R(5, EliminationRole.Entity, "宿州二号", EliminationRole.Year, "2024",
            EliminationRole.Counterparty, "宿州十二号", EliminationRole.Condition, "是",
            EliminationRole.Revenue, "151503.00", EliminationRole.Cost, "147503.00",
            EliminationRole.UnsoldRatio, "10%"));
        rows.Add(R(6, EliminationRole.Entity, "宿州三号", EliminationRole.Year, "2025",
            EliminationRole.Counterparty, "宿州十三号", EliminationRole.Condition, "否",
            EliminationRole.Revenue, "999999", EliminationRole.Cost, "1",
            EliminationRole.UnsoldRatio, "50%"));

        MatrixOptions opt = new MatrixOptions();
        opt.SourceSheetName = "内部交易明细";
        opt.ConditionValue = "是";
        opt.HasConditionColumn = true;
        opt.TotalRow = false;      // 验证表从第 1 行写表头，简化
        opt.FullColumnRefs = true;
        opt.SourceFirstRow = 3;
        opt.SourceLastRow = 1000;
        opt.StartRow = 1;
        opt.StartColumn = 1;

        EliminationMatrix m = EliminationMatrixBuilder.Build(rows, cols, opt);
        Check("构建：2 组 / 2 年度 / 8 行", m.GroupCount == 2 && m.Years.Count == 2 && m.Rows.Count == 8,
            m.GroupCount + "/" + m.Years.Count + "/" + m.Rows.Count);

        // 写入验证表（手工复刻 BuildMatrixPlan 布局：表头 2 行 + 数据）
        dynamic sheet = FreshSheet(book, "验证-收入抵消");
        // 第 1 层表头：年度（公式锚点 G$2 / J$2 用第 2 行，故这里数据从第 4 行起需保持一致）
        // 为与插件一致，仍按 TotalRow=false → 表头在第 1、2 行，数据从第 3 行起
        sheet.Range["G1"].Value2 = "2025";
        sheet.Range["J1"].Value2 = "2024";
        foreach (MatrixRow row in m.Rows)
        {
            sheet.Range["A" + row.RowNumber].Value2 = row.Entity;
            sheet.Range["B" + row.RowNumber].Value2 = row.Customer;
            foreach (string year in m.Years)
            {
                MatrixColumn d = m.DebitColumn(year), c = m.CreditColumn(year);
                MatrixAmount a;
                if (d != null && row.Debit.TryGetValue(year, out a)) sheet.Range[d.Letter + row.RowNumber].Formula = a.Formula;
                if (c != null && row.Credit.TryGetValue(year, out a)) sheet.Range[c.Letter + row.RowNumber].Formula = a.Formula;
            }
        }

        app.Calculate();
        System.Threading.Thread.Sleep(500);

        // 取值（WPS 计算后）
        Func<string, double> val = delegate (string addr)
        {
            try
            {
                object v = sheet.Range[addr].Value2;
                if (v == null) return double.NaN;
                return Convert.ToDouble(v, CultureInfo.InvariantCulture);
            }
            catch { return double.NaN; }
        };
        Func<string, string> txt = delegate (string addr)
        {
            try { return Convert.ToString(sheet.Range[addr].Text); } catch { return ""; }
        };

        // 组1 起始行：TotalRow=false → 表头 2 行 → 数据从第 3 行起；年度锚点为 G$1 / J$1
        // 但插件公式里的年度锚点用的是 firstHeaderRow（= StartRow + TotalRow?1:0 = 1），
        // 且数据 dataTop = 1 + 2 = 3 ✓ 与上面写入一致
        Console.WriteLine("  组1首行 RowNumber=" + m.Rows[0].RowNumber);
        int t1 = m.Rows[0].RowNumber;   // 组1 首行
        int t2 = t1 + 4;                 // 组2 首行（gap=0）
        Check("宿州一号·2025 收入 = 1,998,483.64", Math.Abs(val("G" + t1) - 1998483.64) < 0.01,
            txt("G" + t1));
        Check("宿州一号·2025 成本 = 2,258,286.52", Math.Abs(val("H" + (t1 + 1)) - 2258286.52) < 0.01,
            txt("H" + (t1 + 1)));
        Check("宿州一号·2024 收入 = 2,217,860.59", Math.Abs(val("J" + t1) - 2217860.59) < 0.01,
            txt("J" + t1));
        Check("宿州一号·2024 成本 = 1,849,290.40", Math.Abs(val("K" + (t1 + 1)) - 1849290.40) < 0.01,
            txt("K" + (t1 + 1)));
        // 存货行：2025 走 SUMIFS(W) = -46492.5225；2024 走 SUMPRODUCT 现算 = (R-C)*20% = 73714.038
        Check("宿州一号·2025 存货 = -46,492.52（SUMIFS 损益列）",
            Math.Abs(val("H" + (t1 + 3)) - (-46492.5225)) < 0.01, txt("H" + (t1 + 3)));
        Check("宿州一号·2024 存货 = 73,714.04（SUMPRODUCT 现算，关键）",
            Math.Abs(val("K" + (t1 + 3)) - 73714.038) < 0.05, txt("K" + (t1 + 3)));
        // 成本还原 =(G-H)-K
        Check("宿州一号·2024 成本还原 = 294,856.15",
            Math.Abs(val("K" + (t1 + 2)) - 294856.15) < 0.05, txt("K" + (t1 + 2)));
        Check("宿州二号·2024 收入 = 151,503.00", Math.Abs(val("J" + t2) - 151503.00) < 0.01,
            txt("J" + t2));
        Check("宿州二号·2024 存货 = 400.00（(151503-147503)*10%）",
            Math.Abs(val("K" + (t2 + 3)) - 400.0) < 0.01, txt("K" + (t2 + 3)));
    }

    // ---------------- 往来抵消 ----------------
    static void VerifyReceivable(dynamic app, dynamic book)
    {
        Console.WriteLine();
        Console.WriteLine("--- 往来抵消（差异表，净额法 精确匹配）---");
        List<EliminationColumn> cols = new List<EliminationColumn>();
        cols.Add(C(1, "A", "截止日", EliminationRole.Deadline));
        cols.Add(C(2, "B", "核算主体", EliminationRole.Entity));
        cols.Add(C(3, "C", "科目", EliminationRole.Account));
        cols.Add(C(4, "D", "对方单位", EliminationRole.Counterparty));
        cols.Add(C(6, "F", "是否合并内关联方", EliminationRole.Condition));
        cols.Add(C(21, "U", "审定数", EliminationRole.Amount));

        string A = "芜湖市繁昌区五华鸡产业发展有限公司";
        string B = "芜湖市繁昌区乡村振兴发展集团有限公司";
        List<EliminationRow> rows = new List<EliminationRow>();
        rows.Add(R(4, EliminationRole.Deadline, "20260630", EliminationRole.Entity, A,
            EliminationRole.Account, "其他应收款", EliminationRole.Counterparty, B,
            EliminationRole.Condition, "是", EliminationRole.Amount, "65950000"));
        rows.Add(R(5, EliminationRole.Deadline, "20260630", EliminationRole.Entity, B,
            EliminationRole.Account, "其他应收款", EliminationRole.Counterparty, A,
            EliminationRole.Condition, "是", EliminationRole.Amount, "50000"));
        rows.Add(R(6, EliminationRole.Deadline, "20260630", EliminationRole.Entity, B,
            EliminationRole.Account, "其他应付款", EliminationRole.Counterparty, A,
            EliminationRole.Condition, "是", EliminationRole.Amount, "66000000"));
        rows.Add(R(7, EliminationRole.Deadline, "20260630", EliminationRole.Entity, A,
            EliminationRole.Account, "应收账款", EliminationRole.Counterparty, "外部客户甲",
            EliminationRole.Condition, "否", EliminationRole.Amount, "999"));
        rows.Add(R(8, EliminationRole.Deadline, "20260630", EliminationRole.Entity, A,
            EliminationRole.Account, "应收账款", EliminationRole.Counterparty, B,
            EliminationRole.Condition, "是", EliminationRole.Amount, "800000"));
        rows.Add(R(9, EliminationRole.Deadline, "20260630", EliminationRole.Entity, B,
            EliminationRole.Account, "应付账款", EliminationRole.Counterparty, A,
            EliminationRole.Condition, "是", EliminationRole.Amount, "300000"));

        MatrixOptions opt = new MatrixOptions();
        opt.SourceSheetName = "往来明细表";
        opt.ConditionValue = "是";
        opt.HasConditionColumn = true;
        opt.FullColumnRefs = true;
        opt.SourceFirstRow = 4;
        opt.SourceLastRow = 1000;
        opt.StartRow = 1;
        opt.StartColumn = 1;
        // 结果表金额跨表引用差异表，需先指定差异表名（与下方 FreshSheet 名一致）
        opt.DiffSheetName = "验证-往来差异";

        EliminationDiffTable diff = null;
        EliminationMatrix result = null;
        EliminationDiffBuilder.Build(rows, cols, opt, out diff, out result);
        Check("构建：1 主体对 / 1 年度 / 3 行",
            diff != null && diff.GroupCount == 1 && diff.Years.Count == 1 && diff.Rows.Count == 3,
            diff == null ? "null" : diff.GroupCount + "/" + diff.Years.Count + "/" + diff.Rows.Count);

        dynamic sheet = FreshSheet(book, "验证-往来差异");
        for (int c = 0; c < EliminationDiffTable.Headers.Length; c++)
            sheet.Cells[1, c + 1].Value2 = EliminationDiffTable.Headers[c];
        foreach (EliminationDiffRow row in diff.Rows)
        {
            int r = row.RowNumber;   // 表头占第 1 行，数据自第 2 行起
            sheet.Range["A" + r].Value2 = row.Index;
            sheet.Range["B" + r].Value2 = 2026;
            sheet.Range["C" + r].Value2 = row.Creditor;
            sheet.Range["D" + r].Value2 = row.Debtor;
            sheet.Range["E" + r].Value2 = row.AssetAccount;
            sheet.Range["F" + r].Value2 = row.LiabilityAccount;
            // 与插件 BuildDiffPlan 一致：公式为空（科目不存在）时写数值 0
            Put(sheet, "G" + r, MatrixAmount.Cell(row.Asset, false));
            Put(sheet, "H" + r, MatrixAmount.Cell(row.Liability, false));
            sheet.Range["I" + r].Formula = row.Offset.Formula;
            sheet.Range["J" + r].Formula = row.Difference.Formula;
            sheet.Range["K" + r].Value2 = row.Note;
        }
        app.Calculate();
        System.Threading.Thread.Sleep(500);

        Func<string, double> val = delegate (string addr)
        {
            try { object v = sheet.Range[addr].Value2; return v == null ? double.NaN : Convert.ToDouble(v, CultureInfo.InvariantCulture); }
            catch { return double.NaN; }
        };
        Func<string, string> txt = delegate (string addr)
        {
            try { return Convert.ToString(sheet.Range[addr].Text); } catch { return ""; }
        };

        // 行序：应收账款(2) / 其他应收款(3) / 预付账款(4)
        Check("应收账款对：资产端 = 800,000（净额法）", Math.Abs(val("G2") - 800000.0) < 0.01, txt("G2"));
        Check("应收账款对：负债端 = 300,000", Math.Abs(val("H2") - 300000.0) < 0.01, txt("H2"));
        Check("应收账款对：抵消 = MIN = 300,000", Math.Abs(val("I2") - 300000.0) < 0.01, txt("I2"));
        Check("应收账款对：差额 = 500,000", Math.Abs(val("J2") - 500000.0) < 0.01, txt("J2"));
        Check("其他应收/应付：资产端 = 65,950,000（关键：净额法）",
            Math.Abs(val("G3") - 65950000.0) < 0.01, txt("G3"));
        Check("其他应收/应付：负债端 = 65,950,000", Math.Abs(val("H3") - 65950000.0) < 0.01, txt("H3"));
        Check("其他应收/应付：抵消 = 65,950,000", Math.Abs(val("I3") - 65950000.0) < 0.01, txt("I3"));
        Check("其他应收/应付：差额 = 0", Math.Abs(val("J3") - 0.0) < 0.01, txt("J3"));
        Check("预付/合同负债：无往来 → 0", Math.Abs(val("G4")) < 0.01 && Math.Abs(val("H4")) < 0.01,
            txt("G4") + "/" + txt("H4"));

        // ---- 结果表：跨表引用差异表的抵消金额格 ----
        dynamic resultSheet = FreshSheet(book, "验证-往来结果");
        // 简化布局：第 1 行表头 + 第 2 行起数据（不写 SUBTOTAL/双层表头，只验证公式计算）
        resultSheet.Range["A1"].Value2 = "核算主体";
        resultSheet.Range["B1"].Value2 = "对方";
        resultSheet.Range["E1"].Value2 = "会计科目";
        resultSheet.Range["G1"].Value2 = "2026借";
        resultSheet.Range["H1"].Value2 = "2026贷";
        foreach (MatrixRow row in result.Rows)
        {
            int r = row.RowNumber;   // BuildResult 给的行号从 dataTop=4 起；这里平移到第 2 行起
            int rr = r - 2;          // 简化：dataTop=4 → 第 2 行
            resultSheet.Range["A" + rr].Value2 = row.Entity;
            resultSheet.Range["B" + rr].Value2 = row.Customer;
            resultSheet.Range["E" + rr].Value2 = row.Account;
            MatrixAmount debitAmt, creditAmt;
            if (row.Debit.TryGetValue("2026", out debitAmt))
            {
                object cell = MatrixAmount.Cell(debitAmt, false);
                if (cell is string) resultSheet.Range["G" + rr].Formula = (string)cell;
                else resultSheet.Range["G" + rr].Value2 = cell;
            }
            if (row.Credit.TryGetValue("2026", out creditAmt))
            {
                object cell = MatrixAmount.Cell(creditAmt, false);
                if (cell is string) resultSheet.Range["H" + rr].Formula = (string)cell;
                else resultSheet.Range["H" + rr].Value2 = cell;
            }
        }
        app.Calculate();
        System.Threading.Thread.Sleep(500);
        // 结果表金额应 = 差异表抵消金额（跨表引用生效，不再错乱）
        // 行序（dataTop=4，平移 -2 → 第 2 行起）：应收账款借 G2 / 贷 H3；其他应收借 G4 / 贷 H5
        Check("结果表·应收账款对 借 = 300,000（跨表引用差异表 I2）",
            Math.Abs(val2(resultSheet, "G2") - 300000.0) < 0.01, txt2(resultSheet, "G2"));
        Check("结果表·应收账款对 贷 = 300,000",
            Math.Abs(val2(resultSheet, "H3") - 300000.0) < 0.01, txt2(resultSheet, "H3"));
        Check("结果表·其他应收对 借 = 65,950,000（跨表引用差异表 I3，关键）",
            Math.Abs(val2(resultSheet, "G4") - 65950000.0) < 0.01, txt2(resultSheet, "G4"));
        Check("结果表·其他应收对 贷 = 65,950,000",
            Math.Abs(val2(resultSheet, "H5") - 65950000.0) < 0.01, txt2(resultSheet, "H5"));
    }

    private static double val2(dynamic sheet, string addr)
    {
        try { object v = sheet.Range[addr].Value2; return v == null ? double.NaN : Convert.ToDouble(v, CultureInfo.InvariantCulture); }
        catch { return double.NaN; }
    }
    private static string txt2(dynamic sheet, string addr)
    {
        try { return Convert.ToString(sheet.Range[addr].Text); } catch { return ""; }
    }
}
