using System;
using System.Collections.Generic;
using Excel_Iterate;

namespace Excel_Iterate.TestHarness
{
    // =====================================================================
    // ErrorSourceProbe — 错误来源追溯引擎离线自测（不连 WPS 表格，不需要 COM）
    //
    // 三部分断言：
    //   A. 文本层：FormulaRefParser 的错误字面量纯函数（剔除双引号串）
    //   B. 值层：  ErrorSourceEngine.TryReadError 错误码/文本解码
    //   C. 遍历层：注入假 CellReader 验证 TraceCore 的 BFS 判据
    //      （用户场景 B4 根因、防环、深度上限、取消、自算错误、剪枝、外部未打开）
    //
    // 编译：build.py（/r:Excel_Iterate.dll）
    // 运行：build/ErrorSourceProbe.exe  退出码 0=全 PASS，1=有 FAIL
    // =====================================================================

    internal static class ErrorSourceProbe
    {
        static int _pass = 0, _fail = 0;

        static void Assert(string name, bool cond, string detail)
        {
            if (cond) { _pass++; Console.WriteLine("  [PASS] " + name); }
            else { _fail++; Console.WriteLine("  [FAIL] " + name + " — " + detail); }
        }

        // ---- 假 CellReader：字典 "sheet|addr" → (formula, errorText) ----

        class FakeBook
        {
            public Dictionary<string, string[]> Cells =
                new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);

            /// <summary>登记一格：formula 公式原文（无公式传 ""），err 错误文本（无错误传 null）。</summary>
            public void Set(string sheet, string addr, string formula, string err)
            {
                Cells[sheet + "|" + addr] = new[] { formula, err };
            }

            public CellReadResult Read(string wbPath, string wbName, string sheet, string addr)
            {
                CellReadResult r = new CellReadResult();
                string[] v;
                if (!Cells.TryGetValue(sheet + "|" + addr, out v))
                {
                    // 未登记 = 空白常量格（定位成功、无公式、无错误）
                    r.Found = true;
                    return r;
                }
                r.Found = true;
                r.Formula = v[0] ?? "";
                r.HasErrorValue = v[1] != null;
                r.ErrorText = v[1] ?? "";
                return r;
            }
        }

        static ErrorTraceResult Run(FakeBook book, string rootAddr, ErrorTraceOptions opt)
        {
            return ErrorSourceEngine.TraceCore("C:\\T\\wb.xlsx", "wb.xlsx", "Sheet1",
                new List<string> { rootAddr }, book.Read, opt, null);
        }

        [STAThread]
        static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("=== Excel-iterate ErrorSourceProbe — 错误来源追溯离线自测 ===");
            Console.WriteLine();

            // ===== A. 文本层：错误字面量纯函数 =====
            Console.WriteLine("--- A. 文本层：StripStringLiterals / ContainsErrorLiteral / FirstErrorLiteral ---");
            Assert("含 '#REF!'! 跨表引用 → 命中",
                FormulaRefParser.ContainsErrorLiteral("=D12+'#REF!'!I22"), "未命中");
            Assert("FirstErrorLiteral 返回 #REF!",
                FormulaRefParser.FirstErrorLiteral("=D12+'#REF!'!I22") == "#REF!",
                "实际 " + (FormulaRefParser.FirstErrorLiteral("=D12+'#REF!'!I22") ?? "null"));
            Assert("双引号串内的 #REF! 不算（=\"提示：#REF!\"）",
                !FormulaRefParser.ContainsErrorLiteral("=\"提示：#REF!\""), "误判");
            Assert("普通公式 =A1+A2 不含错误字面量",
                !FormulaRefParser.ContainsErrorLiteral("=A1+A2"), "误判");
            Assert("StripStringLiterals 转义引号 =\"a\"\"b\"+A1 → =+A1",
                FormulaRefParser.StripStringLiterals("=\"a\"\"b\"+A1") == "=+A1",
                "实际 " + FormulaRefParser.StripStringLiterals("=\"a\"\"b\"+A1"));
            Assert("小写 #ref! 忽略大小写命中",
                FormulaRefParser.ContainsErrorLiteral("=#ref!"), "未命中");
            Console.WriteLine();

            // ===== B. 值层：错误码解码 =====
            Console.WriteLine("--- B. 值层：TryReadError ---");
            Assert("TryReadError(2042) = #N/A",
                ErrorSourceEngine.TryReadError(2042) == "#N/A",
                "实际 " + (ErrorSourceEngine.TryReadError(2042) ?? "null"));
            Assert("TryReadError(0x800A07FA) = #N/A（HRESULT 高位剥离）",
                ErrorSourceEngine.TryReadError(unchecked((int)0x800A07FA)) == "#N/A",
                "实际 " + (ErrorSourceEngine.TryReadError(unchecked((int)0x800A07FA)) ?? "null"));
            Assert("TryReadError(2023) = #REF!",
                ErrorSourceEngine.TryReadError(2023) == "#REF!",
                "实际 " + (ErrorSourceEngine.TryReadError(2023) ?? "null"));
            Assert("TryReadError(\"#DIV/0!\") = #DIV/0!（Text 形态）",
                ErrorSourceEngine.TryReadError("#DIV/0!") == "#DIV/0!",
                "实际 " + (ErrorSourceEngine.TryReadError("#DIV/0!") ?? "null"));
            Assert("TryReadError(42) = null（普通数字）",
                ErrorSourceEngine.TryReadError(42) == null, "误判");
            Assert("TryReadError(\"hello\") = null",
                ErrorSourceEngine.TryReadError("hello") == null, "误判");
            Assert("TryReadError(null) = null",
                ErrorSourceEngine.TryReadError(null) == null, "误判");
            Console.WriteLine();

            // ===== C1. 用户场景：A2 错误的根因是 B4 =====
            Console.WriteLine("--- C1. 用户场景：=A1+A2+A3 / =B3+B4 / =D12+'#REF!'!I22 ---");
            FakeBook b1 = new FakeBook();
            b1.Set("Sheet1", "A2", "=A1+A2+A3", "#REF!");
            b1.Set("Sheet1", "A1", "", null);
            b1.Set("Sheet1", "A3", "=B3+B4", "#REF!");
            b1.Set("Sheet1", "B3", "", null);
            b1.Set("Sheet1", "B4", "=D12+'#REF!'!I22", "#REF!");
            b1.Set("Sheet1", "D12", "", null);
            ErrorTraceResult r1 = Run(b1, "A2", null);
            Assert("C1 Success", r1.Success, r1.ErrorMessage);
            Assert("C1 恰好 1 个错误源", r1.Sources.Count == 1, "实际 " + r1.Sources.Count);
            if (r1.Sources.Count == 1)
            {
                ErrorSourceItem s = r1.Sources[0];
                Assert("C1 错误源是 B4", s.CellAddressA1 == "B4", "实际 " + s.CellAddressA1);
                Assert("C1 原因=LiteralInFormula", s.Reason == ErrorSourceReason.LiteralInFormula,
                    "实际 " + s.Reason);
                Assert("C1 Depth=2", s.Depth == 2, "实际 " + s.Depth);
                Assert("C1 PathText=A2 ← A3 ← B4", s.PathText == "A2 ← A3 ← B4",
                    "实际 " + s.PathText);
            }
            Console.WriteLine();

            // ===== C2. 自算错误：=1/0 无上游引用 =====
            Console.WriteLine("--- C2. 自算错误 =1/0 → SelfComputed ---");
            FakeBook b2 = new FakeBook();
            b2.Set("Sheet1", "X1", "=1/0", "#DIV/0!");
            ErrorTraceResult r2 = Run(b2, "X1", null);
            Assert("C2 恰好 1 个错误源", r2.Sources.Count == 1, "实际 " + r2.Sources.Count);
            if (r2.Sources.Count == 1)
            {
                Assert("C2 原因=SelfComputed", r2.Sources[0].Reason == ErrorSourceReason.SelfComputed,
                    "实际 " + r2.Sources[0].Reason);
                Assert("C2 Depth=0（起点自身）", r2.Sources[0].Depth == 0,
                    "实际 " + r2.Sources[0].Depth);
            }
            Console.WriteLine();

            // ===== C3. 剪枝：正常上游不入队，错误链中点的自算错误 =====
            Console.WriteLine("--- C3. 剪枝 + 链中自算错误 ---");
            FakeBook b3 = new FakeBook();
            b3.Set("Sheet1", "Q1", "=Q2+Q3", "#VALUE!");
            b3.Set("Sheet1", "Q2", "=Q4", "#VALUE!");   // 错误，但上游 Q4 正常 → SelfComputed
            b3.Set("Sheet1", "Q3", "", null);           // 正常 → 剪枝
            b3.Set("Sheet1", "Q4", "", null);
            ErrorTraceResult r3 = Run(b3, "Q1", null);
            Assert("C3 恰好 1 个错误源", r3.Sources.Count == 1, "实际 " + r3.Sources.Count);
            if (r3.Sources.Count == 1)
            {
                Assert("C3 错误源是 Q2", r3.Sources[0].CellAddressA1 == "Q2",
                    "实际 " + r3.Sources[0].CellAddressA1);
                Assert("C3 原因=SelfComputed", r3.Sources[0].Reason == ErrorSourceReason.SelfComputed,
                    "实际 " + r3.Sources[0].Reason);
            }
            Console.WriteLine();

            // ===== C4. 多错误源 =====
            Console.WriteLine("--- C4. 多错误源（字面量 + 自算） ---");
            FakeBook b4 = new FakeBook();
            b4.Set("Sheet1", "M1", "=S1+S2", "#REF!");
            b4.Set("Sheet1", "S1", "='#REF!'", "#REF!");   // 字面量源
            b4.Set("Sheet1", "S2", "=NA()", "#N/A");        // 自算源
            ErrorTraceResult r4 = Run(b4, "M1", null);
            Assert("C4 恰好 2 个错误源", r4.Sources.Count == 2, "实际 " + r4.Sources.Count);
            Console.WriteLine();

            // ===== C5. 防环：P1 ⇄ P2 互相引用均报错 =====
            Console.WriteLine("--- C5. 防环（互相引用不死循环） ---");
            FakeBook b5 = new FakeBook();
            b5.Set("Sheet1", "P1", "=P2", "#REF!");
            b5.Set("Sheet1", "P2", "=P1", "#REF!");
            ErrorTraceResult r5 = Run(b5, "P1", null);
            Assert("C5 Success 且不死循环", r5.Success, r5.ErrorMessage);
            Assert("C5 访问节点=2", r5.NodesVisited == 2, "实际 " + r5.NodesVisited);
            Console.WriteLine();

            // ===== C6. 深度上限：长链截断并兜底收录 =====
            Console.WriteLine("--- C6. MaxDepth=3 截断 ---");
            FakeBook b6 = new FakeBook();
            for (int i = 0; i <= 15; i++)
                b6.Set("Sheet1", "C" + i, i == 15 ? "=0" : "=C" + (i + 1), "#NUM!");
            ErrorTraceOptions o6 = new ErrorTraceOptions();
            o6.MaxDepth = 3;
            ErrorTraceResult r6 = Run(b6, "C0", o6);
            Assert("C6 TruncatedByLimit", r6.TruncatedByLimit, "未截断");
            Assert("C6 恰好 1 个兜底错误源 C3", r6.Sources.Count == 1
                && r6.Sources[0].CellAddressA1 == "C3",
                "实际 " + r6.Sources.Count + " 个"
                + (r6.Sources.Count > 0 ? "，首源 " + r6.Sources[0].CellAddressA1 : ""));
            Console.WriteLine();

            // ===== C7. 取消 =====
            Console.WriteLine("--- C7. 用户取消 ---");
            FakeBook b7 = new FakeBook();
            b7.Set("Sheet1", "Z1", "=Z2", "#REF!");
            b7.Set("Sheet1", "Z2", "=#REF!", "#REF!");
            ErrorTraceResult r7 = ErrorSourceEngine.TraceCore("", "wb.xlsx", "Sheet1",
                new List<string> { "Z1" }, b7.Read, null, () => true);
            Assert("C7 Cancelled=true", r7.Cancelled, "未取消");
            Console.WriteLine();

            // ===== C8. 外部引用定位失败：兜底自算 =====
            Console.WriteLine("--- C8. 上游定位失败（外部簿打不开）→ 兜底 SelfComputed ---");
            FakeBook b8 = new FakeBook();
            b8.Set("Sheet1", "E1", "='D:\\x\\[wb2]Ext'!A1+1", "#REF!");
            // 跨簿目标表用定位失败的 reader（FakeBook 对未登记格返回空白格，不适用本例）
            ErrorTraceResult r8b = ErrorSourceEngine.TraceCore("", "wb.xlsx", "Sheet1",
                new List<string> { "E1" },
                (p, n, s, a) =>
                {
                    if (s == "Sheet1") return b8.Read(p, n, s, a);
                    return new CellReadResult { Found = false, FailNote = "工作簿未打开" };
                }, null, null);
            Assert("C8 恰好 1 个兜底错误源 E1", r8b.Sources.Count == 1
                && r8b.Sources[0].CellAddressA1 == "E1"
                && r8b.Sources[0].Reason == ErrorSourceReason.SelfComputed,
                "实际 " + r8b.Sources.Count + " 个"
                + (r8b.Sources.Count > 0 ? "，首源 " + r8b.Sources[0].CellAddressA1
                    + "/" + r8b.Sources[0].Reason : ""));
            Console.WriteLine();

            // ===== 总结 =====
            Console.WriteLine("============================================");
            Console.WriteLine("=== 结果：" + _pass + " PASS / " + _fail + " FAIL ===");
            Console.WriteLine("============================================");
            return _fail > 0 ? 1 : 0;
        }
    }
}
