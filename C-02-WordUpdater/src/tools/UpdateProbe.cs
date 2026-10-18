using System;
using System.Collections.Generic;

namespace WordUpdater.Tools
{
    // =====================================================================
    // UpdateProbe.exe —— 「更新链接」数值变动检测 + 标黄的离线自检工具
    //
    //   UpdateEngine 的输入是 dynamic（WPS 的 Word 对象模型），所以只要
    //   仿出同样形状的一小组对象，就能在没有 WPS 的机器上把
    //   「收集 → 更新前取值 → Update → 重新取值比对 → 定位 → 打黄」
    //   整条链路跑一遍。
    //
    //   断言写的是"正确行为应该是什么"，不是复述实现——
    //   实现改错了这里会红，实现改对了但写法变了这里不该红。
    //
    //   用法：build\UpdateProbe.exe
    // =====================================================================

    public class UpdateProbe
    {
        static int _pass, _fail;

        static void Check(string name, bool ok, string detail)
        {
            if (ok) { _pass++; Console.WriteLine("  PASS  " + name); }
            else { _fail++; Console.WriteLine("  FAIL  " + name + "  —— " + detail); }
        }

        static void Section(string title)
        {
            Console.WriteLine();
            Console.WriteLine("== " + title + " ==");
        }

        // ---------------------------- 假的 Word 对象模型 ----------------------------
        // dynamic 绑定受可访问性限制：这些类与成员必须全部 public。
        // 另外 C# 索引器默认发出的方法名就叫 Item，所以索引器与 Item(int)
        // 不能在同一类里共存（CS0102）—— 这里拆成两个类分别覆盖两条分支。

        public class FakeCells
        {
            private readonly List<object> _list = new List<object>();
            public int Count { get { return _list.Count; } }
            public void Add(object o) { _list.Add(o); }
            public object this[int i] { get { return _list[i - 1]; } }   // 1 起始
        }

        public class FakeCell
        {
            public int RowIndex;
            public int ColumnIndex;
            public FakeCell(int r, int c) { RowIndex = r; ColumnIndex = c; }
        }

        public class FakeParagraphs
        {
            public int Count;
            public FakeParagraphs(int n) { Count = n; }
        }

        /// <summary>一段 Range：字段代码区与结果区都用它（都带 Text/Start/高亮/定位信息）。</summary>
        public class FakeRange
        {
            public string Text;
            public int Start;
            public int HighlightColorIndex;
            public bool InTable;
            public int Page;
            public int ParagraphIndex;      // 表格外兜底用：前面已有多少段
            public FakeCells Cells;
            public FakeDocument Document;

            public FakeRange(string text, int start) { Text = text; Start = start; }

            public object Information(int type)
            {
                if (type == 12) return InTable;       // wdWithInTable
                if (type == 3) return Page;           // wdActiveEndPageNumber
                return 0;
            }
        }

        public class FakeDocument
        {
            public string Name = "测试用例.docx";
            public FakeDocument() { }
            public FakeRange Range(int start, int end)
            {
                // 返回"从头到 end"的一段，段落数由调用方预设
                FakeRange r = new FakeRange("", start);
                r.Text = "";
                return r;
            }
        }

        public class FakeResult : FakeRange
        {
            public FakeResult(string text, int start) : base(text, start) { }
        }

        // ★ 关键：Result 每次访问都返回【新对象】——
        //   真实的 WPS 里 Update() 之后旧 Range 引用会失效，引擎必须重新取，
        //   复用旧对象会拿到过期文本（甚至 null 绑定错误）。
        public class FakeField
        {
            public int Type = 56;                 // wdFieldLink
            public FakeRange Code;
            public bool ShowCodes;
            public string BeforeText;
            public string AfterText;
            public int PrevHighlight;             // 这段文本【原来】的高亮值
            public int StartPos;
            public bool InTable = true;
            public int Row = 1, Col = 1;
            public int ParagraphIndex = 0;
            public int Page = 1;

            public bool UpdateCalled;
            public object UpdateReturns = true;   // null → 视为成功
            public bool UpdateThrows;
            public string ThrowMessage = "源文件不可用";

            public readonly List<FakeResult> CreatedResults = new List<FakeResult>();

            public FakeField(string before, string after, int start)
            {
                BeforeText = before;
                AfterText = after;
                StartPos = start;
                Code = new FakeRange("LINK Excel.Sheet.12 \"C:\\Book.xlsx\" \"Sheet1!R1C1\"", start);
                Code.HighlightColorIndex = PrevHighlight;
            }

            public FakeResult Result
            {
                get
                {
                    string t = UpdateCalled ? AfterText : BeforeText;
                    FakeResult r = new FakeResult(t, StartPos);
                    r.HighlightColorIndex = PrevHighlight;
                    r.InTable = InTable;
                    r.Page = Page;
                    r.ParagraphIndex = ParagraphIndex;
                    r.Cells = new FakeCells();
                    if (InTable) r.Cells.Add(new FakeCell(Row, Col));
                    CreatedResults.Add(r);
                    return r;
                }
            }

            public FakeResult LastResult
            {
                get
                {
                    if (CreatedResults.Count == 0) return null;
                    return CreatedResults[CreatedResults.Count - 1];
                }
            }

            public object Update()
            {
                UpdateCalled = true;
                if (UpdateThrows) throw new System.Runtime.InteropServices.COMException(ThrowMessage, -2146824090);
                return UpdateReturns;
            }
        }

        public class FakeLinkFormat
        {
            public string SourceFullName;
            public int UpdateCalls;
            public FakeLinkFormat(string src) { SourceFullName = src; }
            public void Update() { UpdateCalls++; }
        }

        public class FakeShape
        {
            public FakeLinkFormat LinkFormat;
            public FakeRange Anchor;
            public FakeRange Range;
            public FakeShape(string src, int start)
            {
                LinkFormat = new FakeLinkFormat(src);
                Anchor = new FakeRange("", start);
                Range = new FakeRange("", start);
            }
        }

        public class FakeCollection<T>
        {
            private readonly List<T> _list = new List<T>();
            public int Count { get { return _list.Count; } }
            public void Add(T item) { _list.Add(item); }
            public object this[int i] { get { return _list[i - 1]; } }   // 1 起始
        }

        public class FakeShapes            // 只提供 Item(i)，覆盖 GetItem 的第二条分支
        {
            private readonly List<object> _list = new List<object>();
            public int Count { get { return _list.Count; } }
            public void Add(object o) { _list.Add(o); }
            public object Item(int i) { return _list[i - 1]; }
        }

        public class FakeSelRange
        {
            public int Start = 0;
            public int End = 500;
            public FakeCollection<object> Fields = new FakeCollection<object>();
            public FakeCollection<object> InlineShapes = new FakeCollection<object>();
        }

        public class FakeSelection
        {
            public int Type = 2;                // wdSelectionNormal
            public FakeSelRange Range = new FakeSelRange();
            public FakeShapes ShapeRange = new FakeShapes();
        }

        public class FakeDocuments
        {
            public int Count = 1;
        }

        public class FakeWindow
        {
            public bool ScrollCalled;
            public void ScrollIntoView(object obj, bool down) { ScrollCalled = true; }
            public void ScrollIntoView(object obj) { ScrollCalled = true; }
        }

        public class FakeApplication
        {
            public FakeDocuments Documents = new FakeDocuments();
            public FakeSelection Selection = new FakeSelection();
            public FakeWindow ActiveWindow = new FakeWindow();
            public object DisplayAlerts = 2;
            public object ScreenUpdating = true;
        }

        // ---------------------------- 场景搭建 ----------------------------

        static FakeField AddField(FakeApplication app, string before, string after, int start)
        {
            FakeField f = new FakeField(before, after, start);
            app.Selection.Range.Fields.Add(f);
            return f;
        }

        static FakeApplication NewApp()
        {
            return new FakeApplication();
        }

        public static int Main()
        {
            Console.WriteLine("UpdateProbe —— 「更新链接」变动检测 + 标黄 离线自检");
            Console.WriteLine("（假对象模型，不需要 WPS）");

            // ---------------------------------------------------------------
            Section("1. 值发生变动 → 记录 + 定位 + 标黄");
            {
                FakeApplication app = NewApp();
                FakeField f = AddField(app, "100", "200", 10);
                f.Row = 2; f.Col = 2;
                RunResult r = UpdateEngine.Execute(app);

                Check("只有 1 条变动", r.Changes.Count == 1, "实际 " + r.Changes.Count);
                Check("更新成功 1 条、失败 0 条", r.Success == 1 && r.Failed == 0,
                    "success=" + r.Success + " failed=" + r.Failed);
                if (r.Changes.Count == 1)
                {
                    LinkChangeInfo c = r.Changes[0];
                    Check("旧值/新值正确", c.OldText == "100" && c.NewText == "200",
                        c.OldText + " → " + c.NewText);
                    Check("定位成『第 2 行第 2 列』", c.Where == "第 2 行第 2 列", c.Where);
                    Check("标黄打在【更新后重新取到】的那个 Result 上",
                        ReferenceEquals(c.FlashRange, f.LastResult), "用的不是最新的 Result");
                    Check("高亮值被设成 wdYellow(7)",
                        f.LastResult != null && f.LastResult.HighlightColorIndex == 7,
                        "实际 " + (f.LastResult == null ? "null" : f.LastResult.HighlightColorIndex.ToString()));
                    Check("记下的原值是 0", c.PrevHighlight == 0, "实际 " + c.PrevHighlight);
                }
                Check("图片/对象计数为 0", r.NoTextCount == 0, "实际 " + r.NoTextCount);
            }

            // ---------------------------------------------------------------
            Section("2. 值没变 → 不记录、不标黄");
            {
                FakeApplication app = NewApp();
                FakeField f = AddField(app, "100", "100", 10);
                RunResult r = UpdateEngine.Execute(app);

                Check("变动清单为空", r.Changes.Count == 0, "实际 " + r.Changes.Count);
                Check("仍算更新成功", r.Success == 1, "实际 " + r.Success);
                bool untouched = true;
                for (int i = 0; i < f.CreatedResults.Count; i++)
                    if (f.CreatedResults[i].HighlightColorIndex != 0) untouched = false;
                Check("高亮始终没被碰过（保持原值 0）", untouched, "有 Result 的高亮被改了");
                Check("图片/对象计数为 0", r.NoTextCount == 0, "实际 " + r.NoTextCount);
            }

            // ---------------------------------------------------------------
            Section("3. Update() 返回 false → 只记失败，不进变动");
            {
                FakeApplication app = NewApp();
                FakeField f = AddField(app, "100", "200", 10);
                f.UpdateReturns = false;
                RunResult r = UpdateEngine.Execute(app);

                Check("失败 1 条、成功 0 条", r.Failed == 1 && r.Success == 0,
                    "success=" + r.Success + " failed=" + r.Failed);
                Check("变动清单为空", r.Changes.Count == 0, "实际 " + r.Changes.Count);
                Check("没有任何标黄", NoHighlight(f), "有 Result 被标黄了");
                Check("图片/对象计数为 0（失败项不该算『无文本』）", r.NoTextCount == 0, "实际 " + r.NoTextCount);
            }

            // ---------------------------------------------------------------
            Section("4. Update() 抛 COM 异常 → 记失败，引擎不崩");
            {
                FakeApplication app = NewApp();
                FakeField f = AddField(app, "100", "200", 10);
                f.UpdateThrows = true;
                RunResult r = UpdateEngine.Execute(app);

                Check("失败 1 条", r.Failed == 1, "实际 " + r.Failed);
                Check("变动清单为空", r.Changes.Count == 0, "实际 " + r.Changes.Count);
                Check("状态仍是 Completed（不是引擎级 Error）",
                    r.Status == RunStatus.Completed, r.Status.ToString());
                Check("没有 ErrorMessage", string.IsNullOrEmpty(r.ErrorMessage), r.ErrorMessage);
            }

            // ---------------------------------------------------------------
            Section("5. 链接图片/形状 → 只算已更新，不参与比对");
            {
                FakeApplication app = NewApp();
                FakeShape sh = new FakeShape("C:\\pic.xlsx", 50);
                app.Selection.Range.InlineShapes.Add(sh);
                RunResult r = UpdateEngine.Execute(app);

                Check("发现 1 个、成功 1 个", r.Found == 1 && r.Success == 1,
                    "found=" + r.Found + " success=" + r.Success);
                Check("LinkFormat.Update 被调用 1 次", sh.LinkFormat.UpdateCalls == 1,
                    "实际 " + sh.LinkFormat.UpdateCalls);
                Check("计入『无文本可比』", r.NoTextCount == 1, "实际 " + r.NoTextCount);
                Check("变动清单为空（图片没有可比对的数字）", r.Changes.Count == 0, "实际 " + r.Changes.Count);
            }

            // ---------------------------------------------------------------
            Section("6. 表格外的链接 → 走正文兜底文案，但仍要标黄");
            {
                FakeApplication app = NewApp();
                FakeField f = AddField(app, "100", "200", 10);
                f.InTable = false;
                f.Page = 3;
                RunResult r = UpdateEngine.Execute(app);

                Check("有 1 条变动", r.Changes.Count == 1, "实际 " + r.Changes.Count);
                if (r.Changes.Count == 1)
                {
                    string w = r.Changes[0].Where;
                    Check("位置文案不是『第 X 行』", w.IndexOf("第", StringComparison.Ordinal) != 0
                        && w.IndexOf("行", StringComparison.Ordinal) < 0, w);
                    Check("位置文案是正文兜底", w.IndexOf("正文", StringComparison.Ordinal) >= 0, w);
                    Check("仍然标黄了", r.Changes[0].FlashRange != null, "FlashRange 为 null");
                }
            }

            // ---------------------------------------------------------------
            Section("7. 原值还原：3 → 黄 7 → 回到 3（不是清零）");
            {
                FakeRange rg = new FakeRange("100", 0);
                rg.HighlightColorIndex = 3;
                int prev = -1;
                bool applied = FlashHighlighter.Apply(rg, out prev);
                Check("打黄成功且原值是 3", applied && prev == 3, "applied=" + applied + " prev=" + prev);
                Check("现在是 wdYellow(7)", rg.HighlightColorIndex == 7, "实际 " + rg.HighlightColorIndex);
                FlashHighlighter.Clear(rg, prev);
                Check("还原后回到 3", rg.HighlightColorIndex == 3, "实际 " + rg.HighlightColorIndex);
            }

            // ---------------------------------------------------------------
            Section("8. ShowCodes（Alt+F9）下黄要打在 Code 上");
            {
                FakeApplication app = NewApp();
                FakeField f = AddField(app, "100", "200", 10);
                f.ShowCodes = true;
                RunResult r = UpdateEngine.Execute(app);

                Check("有 1 条变动", r.Changes.Count == 1, "实际 " + r.Changes.Count);
                if (r.Changes.Count == 1)
                {
                    Check("黄打在 Code 上", ReferenceEquals(r.Changes[0].FlashRange, f.Code),
                        "FlashRange 不是 Code");
                    Check("Code 的高亮是 7", f.Code.HighlightColorIndex == 7,
                        "实际 " + f.Code.HighlightColorIndex);
                    bool resultUntouched = true;
                    for (int i = 0; i < f.CreatedResults.Count; i++)
                        if (f.CreatedResults[i].HighlightColorIndex != 0) resultUntouched = false;
                    Check("Result 没被碰", resultUntouched, "Result 被标黄了");
                }
            }

            // ---------------------------------------------------------------
            Section("9. 更新后必须【重新】取 Result（不能复用旧 Range）");
            {
                FakeApplication app = NewApp();
                FakeField f = AddField(app, "100", "200", 10);
                UpdateEngine.Execute(app);

                Check("Result 被取了至少 3 次（收集 + 更新前 + 更新后）",
                    f.CreatedResults.Count >= 3, "实际 " + f.CreatedResults.Count);
                bool allFresh = true;
                for (int i = 0; i < f.CreatedResults.Count; i++)
                {
                    // 更新前的那几次读到 100，更新后那次读到 200
                    if (f.CreatedResults[i].Text != "100" && f.CreatedResults[i].Text != "200")
                        allFresh = false;
                }
                Check("每次取到的都是当时最新的文本", allFresh, "有取到过期文本");
            }

            // ---------------------------------------------------------------
            Section("10. 混合：2 变 + 1 不变 + 1 失败 + 1 图片");
            {
                FakeApplication app = NewApp();
                FakeField f1 = AddField(app, "100", "200", 10); f1.Row = 1; f1.Col = 1;
                FakeField f2 = AddField(app, "300", "400", 20); f2.Row = 1; f2.Col = 2;
                FakeField f3 = AddField(app, "500", "500", 30);           // 不变
                FakeField f4 = AddField(app, "600", "700", 40);
                f4.UpdateReturns = false;                                  // 失败
                FakeShape sh = new FakeShape("C:\\pic.xlsx", 90);
                app.Selection.Range.InlineShapes.Add(sh);

                RunResult r = UpdateEngine.Execute(app);

                Check("发现 5 个", r.Found == 5, "实际 " + r.Found);
                Check("成功 4 个", r.Success == 4, "实际 " + r.Success);
                Check("失败 1 个", r.Failed == 1, "实际 " + r.Failed);
                Check("Found == Success + Failed", r.Found == r.Success + r.Failed,
                    r.Found + " != " + (r.Success + r.Failed));
                Check("变动 2 条", r.Changes.Count == 2, "实际 " + r.Changes.Count);
                Check("无文本 1 个（只有那张图片）", r.NoTextCount == 1, "实际 " + r.NoTextCount);
                if (r.Changes.Count == 2)
                {
                    Check("变动的是第 1 行第 1 列与第 1 行第 2 列",
                        r.Changes[0].Where == "第 1 行第 1 列" && r.Changes[1].Where == "第 1 行第 2 列",
                        r.Changes[0].Where + " / " + r.Changes[1].Where);
                    Check("失败的那个（600→700）没被记成变动",
                        !HasChange(r, "600") && !HasChange(r, "700"), "失败项混进了变动清单");
                }
                Check("不变的 500 没被记成变动", !HasChange(r, "500"), "500 被误报成变动");
            }

            // ---------------------------------------------------------------
            Section("11. 回归：InlineShapes 与 ShapeRange 同锚点只更新一次");
            {
                FakeApplication app = NewApp();
                FakeShape sh = new FakeShape("C:\\pic.xlsx", 70);
                app.Selection.Range.InlineShapes.Add(sh);
                app.Selection.ShapeRange.Add(sh);       // 同一个对象、同一个锚点
                RunResult r = UpdateEngine.Execute(app);

                Check("只发现 1 个（去重生效）", r.Found == 1, "实际 " + r.Found);
                Check("只更新 1 次", sh.LinkFormat.UpdateCalls == 1, "实际 " + sh.LinkFormat.UpdateCalls);
                Check("成功 1 个", r.Success == 1, "实际 " + r.Success);
            }

            // ---------------------------------------------------------------
            Section("12. FlashSession 多槽还原 + 引用去重保最早原值");
            {
                FakeRange a = new FakeRange("1", 0);
                FakeRange b = new FakeRange("2", 10);
                a.HighlightColorIndex = 0;
                b.HighlightColorIndex = 5;

                FlashSession s = new FlashSession();
                int p1, p2;
                FlashHighlighter.Apply(a, out p1);
                FlashHighlighter.Apply(b, out p2);
                s.Add(a, p1);
                s.Add(b, p2);
                s.Add(a, 999);            // 重复登记同一个 Range：应被忽略
                Check("登记 2 条（重复的那条被忽略）", s.Count == 2, "实际 " + s.Count);
                Check("两条都已是黄色", a.HighlightColorIndex == 7 && b.HighlightColorIndex == 7,
                    a.HighlightColorIndex + "/" + b.HighlightColorIndex);

                s.ClearNow();
                Check("第一条还原成 0", a.HighlightColorIndex == 0, "实际 " + a.HighlightColorIndex);
                Check("第二条还原成【原值 5】而不是 0", b.HighlightColorIndex == 5,
                    "实际 " + b.HighlightColorIndex);
                Check("清空后没有悬挂条目", s.Count == 0 && !s.HasPending, "还有残留");

                // 同一 Range 被两套机制先后打黄：后写者的 prev 是 7，
                // 去重后必须保留最早的原值，否则还原成 7 = 黄永久残留
                FakeRange c = new FakeRange("3", 0);
                c.HighlightColorIndex = 2;
                FlashSession s2 = new FlashSession();
                int q1, q2;
                FlashHighlighter.Apply(c, out q1);      // 先来者：原值 2
                FlashHighlighter.Apply(c, out q2);      // 后来者：读到的是 7
                s2.Add(c, q1);
                s2.Add(c, q2);
                s2.ClearNow();
                Check("重复打黄后仍还原到【最早的原值 2】", c.HighlightColorIndex == 2,
                    "实际 " + c.HighlightColorIndex + "（若是 7 说明黄永久残留）");
            }

            // ---------------------------------------------------------------
            Section("13. 没有选区 / 没有链接");
            {
                FakeApplication app = NewApp();
                app.Selection.Type = 1;                 // wdSelectionIP
                app.Selection.Range.Start = 10;
                app.Selection.Range.End = 10;
                RunResult r = UpdateEngine.Execute(app);
                Check("插入点状态 → NoSelection", r.Status == RunStatus.NoSelection, r.Status.ToString());
                Check("变动清单为空", r.Changes.Count == 0, "实际 " + r.Changes.Count);

                FakeApplication app2 = NewApp();        // 有选区但没链接
                RunResult r2 = UpdateEngine.Execute(app2);
                Check("空选区 → NoLinks", r2.Status == RunStatus.NoLinks, r2.Status.ToString());
                Check("变动清单为空", r2.Changes.Count == 0, "实际 " + r2.Changes.Count);
            }

            // ---------------------------------------------------------------
            Section("14. 归一化：尾部的单元格结束符不算变动");
            {
                FakeApplication app = NewApp();
                FakeField f = AddField(app, "100\r\a", "100", 10);   // 只差结束符
                RunResult r = UpdateEngine.Execute(app);
                Check("判为【无变动】", r.Changes.Count == 0, "实际 " + r.Changes.Count);
                Check("仍算更新成功", r.Success == 1, "实际 " + r.Success);

                FakeApplication app2 = NewApp();
                FakeField f2 = AddField(app2, "100 ", "200", 10);    // 真变了（空格不掩值变）
                RunResult r2 = UpdateEngine.Execute(app2);
                Check("值真的变了仍然报变动", r2.Changes.Count == 1, "实际 " + r2.Changes.Count);
                if (r2.Changes.Count == 1)
                    Check("旧值已去掉尾随空格", r2.Changes[0].OldText == "100", "[" + r2.Changes[0].OldText + "]");
            }

            // ---------------------------------------------------------------
            Section("15. 弹窗前不抢选区：ScrollIntoView 失败也要静默");
            {
                // 没有 ActiveWindow 的应用：滚动必须静默失败，绝不能抛出去
                FakeApplication app = NewApp();
                FakeRange rg = new FakeRange("x", 0);
                bool scrolled = FlashHighlighter.ScrollIntoView(null, rg);
                Check("app 为 null 时静默返回 false", !scrolled, "返回了 true");

                FakeApplication app2 = NewApp();
                bool scrolled2 = FlashHighlighter.ScrollIntoView(app2, rg);
                Check("有窗口时能滚动", scrolled2 && app2.ActiveWindow.ScrollCalled, "没滚成");
            }

            // ---------------------------------------------------------------
            Console.WriteLine();
            Console.WriteLine("--------------------------------------------");
            Console.WriteLine(string.Format("PASS {0}    FAIL {1}", _pass, _fail));
            return _fail == 0 ? 0 : 1;
        }

        static bool NoHighlight(FakeField f)
        {
            for (int i = 0; i < f.CreatedResults.Count; i++)
                if (f.CreatedResults[i].HighlightColorIndex != f.PrevHighlight) return false;
            return f.Code.HighlightColorIndex == f.PrevHighlight;
        }

        static bool HasChange(RunResult r, string text)
        {
            for (int i = 0; i < r.Changes.Count; i++)
            {
                if (r.Changes[i].OldText == text || r.Changes[i].NewText == text) return true;
            }
            return false;
        }
    }
}
