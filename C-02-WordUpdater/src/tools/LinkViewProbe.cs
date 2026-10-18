using System;
using System.Collections.Generic;
using System.IO;

namespace WordUpdater.Tools
{
    // =====================================================================
    // LinkViewProbe.exe —— 「查看链接」引擎的离线自检工具
    //
    //   LinkViewEngine 的输入是 dynamic（WPS 的 Word 对象模型），
    //   所以只要仿出同样形状的一小组对象，就能在没有 WPS 的机器上把
    //   「扫描 → 聚合 → 更新全部 → 中断全部」整条链路跑一遍。
    //
    //   用法：build\LinkViewProbe.exe
    // =====================================================================

    public class LinkViewProbe
    {
        // ---------------------------- 假的 Word 对象模型 ----------------------------
        // 注意：dynamic 绑定受可访问性限制，这些类必须 public，
        //       否则引擎里 app.ActiveDocument 会静默取不到文档。

        public class FakeCode
        {
            public string Text;
            public int Start;
            public bool SelectCalled;
            public int HighlightColorIndex;
            public FakeCode(string text, int start) { Text = text; Start = start; }
            public void Select() { SelectCalled = true; }
        }

        public class FakeField
        {
            public int Type = 56;
            public FakeCode Code;
            public FakeCode Result;   // 字段显示出来的内容区（标黄目标；null = 走 Code 兜底）
            public bool UpdateCalled;
            public bool UnlinkCalled;
            public bool UpdateReturnsTrue = true;

            public FakeField(string code, int start)
            {
                Code = new FakeCode(code, start);
            }

            public object Update()
            {
                UpdateCalled = true;
                return UpdateReturnsTrue;
            }

            public void Unlink() { UnlinkCalled = true; }
        }

        public class FakeLinkFormat
        {
            public string SourceFullName;
            public int AutoUpdate;
            public int UpdateCalls;
            public int BreakCalls;
            public FakeLinkFormat(string src, int autoUpdate)
            {
                SourceFullName = src;
                AutoUpdate = autoUpdate;
            }
            public void Update() { UpdateCalls++; }
            public void BreakLink() { BreakCalls++; }
        }

        public class FakeAnchor
        {
            public int Start;
            public bool SelectCalled;
            public int HighlightColorIndex;
            public FakeAnchor(int start) { Start = start; }
            public void Select() { SelectCalled = true; }
        }

        /// <summary>内嵌对象专用的"可选中 Range"（浮动形状没有这个属性）。</summary>
        public class FakeRange2
        {
            public bool SelectCalled;
            public int HighlightColorIndex;
            public void Select() { SelectCalled = true; }
        }

        public class FakeShape
        {
            public FakeLinkFormat LinkFormat;
            public FakeAnchor Anchor;
            public FakeRange2 Range;   // null = 浮动形状（跳转走 Anchor 分支）
            public FakeShape(string src, int start, int autoUpdate)
            {
                LinkFormat = new FakeLinkFormat(src, autoUpdate);
                Anchor = new FakeAnchor(start);
            }
        }

        // 集合：GetItem 先试 coll[i]（索引器），失败再试 Item(i)。
        // 本类只提供索引器（注意：C# 索引器默认发出的方法名就叫 Item，
        // 这里不能再手写同名方法），Item(i) 那条分支交给下面的 FakeShapes 覆盖。
        public class FakeCollection<T>
        {
            private readonly List<T> _list = new List<T>();
            public int Count { get { return _list.Count; } }
            public void Add(T item) { _list.Add(item); }
            public object this[int i] { get { return _list[i - 1]; } }   // 1 起始
        }

        public class FakeRange
        {
            public string Name;
            public FakeCollection<object> Fields;
            public FakeCollection<object> InlineShapes;
            public object NextStoryRange { get { return null; } }
            public FakeRange(string name)
            {
                Name = name;
                Fields = new FakeCollection<object>();
                InlineShapes = new FakeCollection<object>();
            }
        }

        public class FakeShapes
        {
            private readonly List<object> _list = new List<object>();
            public int Count { get { return _list.Count; } }
            public void Add(object o) { _list.Add(o); }
            public object Item(int i) { return _list[i - 1]; }
        }

        public class FakeDocument
        {
            public string Name = "测试用例.docx";
            public FakeCollection<object> StoryRanges;
            public FakeShapes Shapes;
            public FakeRange Main;
            public FakeRange Header;
            public FakeDocument()
            {
                Main = new FakeRange("正文");
                Header = new FakeRange("页眉");
                StoryRanges = new FakeCollection<object>();
                StoryRanges.Add(Main);
                StoryRanges.Add(Header);
                Shapes = new FakeShapes();
            }
        }

        public class FakeApplication
        {
            public FakeDocument ActiveDocument;
            public object DisplayAlerts;
            public object ScreenUpdating;
            public int ActivateCalls;
            public FakeApplication(FakeDocument doc) { ActiveDocument = doc; }
            public void Activate() { ActivateCalls++; }
        }

        // ---------------------------- 主流程 ----------------------------

        public static int Main(string[] args)
        {
            int failed = 0;
            string missingFile = Path.Combine(Path.GetTempPath(), "确实不存在的文件.xlsx");
            string existingFile = Path.Combine(Path.GetTempPath(), "LinkViewProbe_存在的源.xlsx");
            try { File.WriteAllText(existingFile, "xlsx"); }
            catch (Exception ex)
            {
                Console.WriteLine("[准备失败] 无法创建临时文件：" + ex.Message);
                return 2;
            }

            FakeApplication app = BuildDoc(existingFile, missingFile);

            Console.WriteLine("================ 场景 1：扫描并聚合 ================");
            LinkViewResult r = LinkViewEngine.Scan(app);
            Console.WriteLine("状态：" + r.Status + "  文档：" + r.DocumentName
                + "  链接总数：" + r.TotalLinks + "  源文件数：" + r.Sources.Count);
            foreach (LinkSourceInfo si in r.Sources)
            {
                Console.WriteLine(string.Format(
                    "  [{0}] 工作表={1} 形式={2} 更新={3} 链接={4} 状态={5}",
                    si.WorkbookName, si.SheetText, si.KindText, si.UpdateText,
                    si.LinkCount, si.StatusText));
                Console.WriteLine("        路径：" + si.FullPath);
                foreach (LinkItemInfo it in si.Items)
                {
                    Console.WriteLine(string.Format(
                        "        #{0} {1} / {2} / {3} / {4}",
                        it.Index, it.ItemKind, it.SheetName, it.RangeText, it.UpdateMode));
                }
            }

            Check(ref failed, "扫描状态应为 Completed", r.Status == LinkViewStatus.Completed);
            Check(ref failed, "应扫到 2 个源文件", r.Sources.Count == 2);
            Check(ref failed, "应去重为 5 条链接（正文 3 字段 + 页眉 1 字段 + 图 1 对象）", r.TotalLinks == 5);

            LinkSourceInfo known = Find(r.Sources, existingFile);
            LinkSourceInfo gone = Find(r.Sources, missingFile);
            Check(ref failed, "存在的源文件状态=正常", known != null && known.StatusText == "正常");
            Check(ref failed, "缺失的源文件状态=源文件缺失", gone != null && gone.StatusText == "源文件缺失");
            Check(ref failed, "工作簿名取的是文件名", known != null && known.WorkbookName == Path.GetFileName(existingFile));
            Check(ref failed, "工作表应聚合出 Sheet1 / Sheet2 / 总表",
                known != null && known.Sheets.Count == 3 && known.SheetText.IndexOf("Sheet1") >= 0);
            Check(ref failed, "更新方式应识别为混合", known != null && known.UpdateText.IndexOf("混合") >= 0);
            Check(ref failed, "区域应解析出 Sheet1!R1C1:R4C4 这类写法",
                known != null && known.Items.Count > 0 && known.Items[0].RangeText.IndexOf("R") == 0);

            Console.WriteLine();
            Console.WriteLine("================ 场景 2：更新所有链接 ================");
            LinkActionResult up = LinkViewEngine.UpdateAll(app);
            Console.WriteLine(string.Format("执行={0} 共={1} 成功={2} 失败={3} 失败明细={4}",
                up.Executed, up.Total, up.Success, up.Failed, up.Failures.Count));
            Check(ref failed, "更新应全部成功", up.Executed && up.Total == 5 && up.Success == 5 && up.Failed == 0);
            Check(ref failed, "运行后主机状态（告警/刷新）应被还原",
                app.ScreenUpdating == null && app.DisplayAlerts == null);

            FakeField f1 = (FakeField)((FakeCollection<object>)app.ActiveDocument.Main.Fields)[1];
            Check(ref failed, "字段的 Update() 被调用", f1.UpdateCalled);

            Console.WriteLine();
            Console.WriteLine("================ 场景 3：中断所有链接 ================");
            LinkActionResult br = LinkViewEngine.BreakAll(app);
            Console.WriteLine(string.Format("执行={0} 共={1} 成功={2} 失败={3}",
                br.Executed, br.Total, br.Success, br.Failed));
            Check(ref failed, "中断应全部成功", br.Executed && br.Total == 5 && br.Success == 5 && br.Failed == 0);
            Check(ref failed, "字段的 Unlink() 被调用", f1.UnlinkCalled);
            FakeShape sh = (FakeShape)app.ActiveDocument.Shapes.Item(1);
            Check(ref failed, "对象的 BreakLink() 被调用", sh.LinkFormat.BreakCalls == 1);

            Console.WriteLine();
            Console.WriteLine("================ 场景 4：空文档 ================");
            FakeApplication empty = new FakeApplication(new FakeDocument());
            LinkViewResult er = LinkViewEngine.Scan(empty);
            Console.WriteLine("状态：" + er.Status);
            Check(ref failed, "无链接文档应返回 NoLinks", er.Status == LinkViewStatus.NoLinks);
            Check(ref failed, "空文档执行更新不应报错", LinkViewEngine.UpdateAll(empty).Executed);

            Console.WriteLine();
            Console.WriteLine("================ 场景 5：跳转到该字段 ================");
            // 单独再搭一个小文档：1 字段 + 1 内嵌 OLE（带 Range），覆盖三个跳转分支
            FakeDocument doc2 = new FakeDocument();
            FakeField ff = new FakeField(
                " LINK Excel.Sheet.12 \"" + existingFile + "\" \"S1!R1C1\" \\t ", 50);
            ff.Result = new FakeCode("字段显示出来的结果", 55);
            doc2.Main.Fields.Add(ff);
            FakeShape ole = new FakeShape(existingFile, 80, 1);
            ole.Range = new FakeRange2();
            doc2.Main.InlineShapes.Add(ole);
            FakeApplication app2 = new FakeApplication(doc2);

            LinkViewResult sr = LinkViewEngine.Scan(app2);
            Check(ref failed, "跳转测试文档应扫到 2 条链接",
                sr.Status == LinkViewStatus.Completed && sr.TotalLinks == 2);

            LinkItemInfo fieldItem = null;
            LinkItemInfo oleItem = null;
            foreach (LinkSourceInfo si2 in sr.Sources)
            {
                foreach (LinkItemInfo it in si2.Items)
                {
                    if (it.KindCode == "OLE") oleItem = it;
                    else fieldItem = it;
                }
            }
            Check(ref failed, "应分别找到字段条目与对象条目", fieldItem != null && oleItem != null);

            string e1 = LinkViewEngine.JumpToItem(app2, fieldItem);
            Check(ref failed, "跳转字段应成功", e1 == null);
            Check(ref failed, "字段的 Code.Select() 被调用", ff.Code.SelectCalled);

            string e2 = LinkViewEngine.JumpToItem(app2, oleItem);
            Check(ref failed, "跳转内嵌对象应成功", e2 == null);
            Check(ref failed, "内嵌对象的 Range.Select() 被调用（不碰 Anchor）",
                ole.Range.SelectCalled && !ole.Anchor.SelectCalled);
            Check(ref failed, "两次成功跳转后 Activate() 被调两次", app2.ActivateCalls == 2);

            // 浮动形状（没有 Range）应落到 Anchor 分支 —— 复用场景 1 的文档
            LinkViewResult sr1 = LinkViewEngine.Scan(app);
            LinkItemInfo floatItem = null;
            foreach (LinkSourceInfo si3 in sr1.Sources)
            {
                foreach (LinkItemInfo it in si3.Items)
                {
                    if (it.KindCode == "OLE") floatItem = it;
                }
            }
            Check(ref failed, "场景1 文档里应找到浮动形状条目", floatItem != null);
            FakeShape fl = (FakeShape)app.ActiveDocument.Shapes.Item(1);
            string e3 = LinkViewEngine.JumpToItem(app, floatItem);
            Check(ref failed, "跳转浮动形状应成功", e3 == null);
            Check(ref failed, "浮动形状的 Anchor.Select() 被调用", fl.Anchor.SelectCalled);

            // 标黄提示：JumpToItemEx 应把【可见的结果区 Result】标黄（7）并交出
            // 原高亮值，ClearFlash 按原值还原 —— 先故意设一个非 0 的原值验证
            // "还原不是清零"。
            // （v1.3.3~v1.3.5 的 bug：黄打在 Code 上——字段以结果显示时 Code 是
            //   隐藏文字，黄了也看不见。修复后必须打 Result，Code 全程不许被碰。）
            ff.Result.HighlightColorIndex = 3;
            object flash; int prev;
            string e3b = LinkViewEngine.JumpToItemEx(app2, fieldItem, out flash, out prev);
            Check(ref failed, "Ex 跳转字段应成功", e3b == null);
            Check(ref failed, "Ex 应交出被标黄的 Range（字段的 Result，不是 Code）",
                flash != null && ReferenceEquals(flash, ff.Result));
            Check(ref failed, "Ex 应记下 Result 的原高亮值 3", prev == 3);
            Check(ref failed, "Ex 跳转后 Result 应被标黄（wdYellow=7）",
                ff.Result.HighlightColorIndex == 7);
            Check(ref failed, "Code（隐藏文字）不应被碰",
                ff.Code.HighlightColorIndex == 0);
            LinkViewEngine.ClearFlash(flash, prev);
            Check(ref failed, "ClearFlash 应把 Result 高亮还原为原值 3",
                ff.Result.HighlightColorIndex == 3);
            LinkViewEngine.ClearFlash(null, 0);   // null 容错，不抛即通过
            Check(ref failed, "ClearFlash(null) 应静默通过", true);

            // 兜底：字段拿不到 Result（异常/为空）时退回 Code，黄至少还在
            ff.Result = null;
            string e3c = LinkViewEngine.JumpToItemEx(app2, fieldItem, out flash, out prev);
            Check(ref failed, "Ex 兜底跳转应成功", e3c == null);
            Check(ref failed, "拿不到 Result 时应退回标黄 Code",
                flash != null && ReferenceEquals(flash, ff.Code));
            Check(ref failed, "兜底时 Code 应被标黄（wdYellow=7）",
                ff.Code.HighlightColorIndex == 7);
            LinkViewEngine.ClearFlash(flash, prev);
            Check(ref failed, "兜底还原后 Code 高亮应回到 0",
                ff.Code.HighlightColorIndex == 0);

            // 过期检测：序号越界 / 位置内容被改，都应返回错误串而不是跳错地方
            LinkItemInfo stale = new LinkItemInfo();
            stale.Index = 99;
            stale.KindCode = "LINK";
            string e4 = LinkViewEngine.JumpToItem(app2, stale);
            Check(ref failed, "序号越界应返回错误提示", e4 != null && e4.Length > 0);

            ff.Code.Text = " LINK 已被改掉的代码 ";
            string e5 = LinkViewEngine.JumpToItem(app2, fieldItem);
            Check(ref failed, "字段被改后应检测到列表过期", e5 != null && e5.Length > 0);

            try { File.Delete(existingFile); }
            catch { }

            Console.WriteLine();
            Console.WriteLine(failed == 0 ? "全部自检通过。"
                : string.Format("有 {0} 项自检未通过。", failed));
            return failed == 0 ? 0 : 1;
        }

        // ---------------------------- 构造用例 ----------------------------

        private static FakeApplication BuildDoc(string existingFile, string missingFile)
        {
            FakeDocument doc = new FakeDocument();

            // 正文：3 条 LINK 字段 + 1 条重复（同位置同代码，应被去重）
            doc.Main.Fields.Add(new FakeField(
                " LINK Excel.Sheet.12 \"" + existingFile + "\" \"Sheet1!R1C1:R4C4\" \\t \\a ", 100));
            doc.Main.Fields.Add(new FakeField(
                " LINK Excel.Sheet.12 \"" + existingFile + "\" \"Sheet2!R2C2\" \\t ", 160));
            doc.Main.Fields.Add(new FakeField(
                " LINK Excel.Sheet.12 \"" + missingFile + "\" \"Sheet1!A1:D4\" \\t ", 200));
            doc.Main.Fields.Add(new FakeField(   // 与上一条完全重复 -> 去重
                " LINK Excel.Sheet.12 \"" + missingFile + "\" \"Sheet1!A1:D4\" \\t ", 200));

            // 页眉：DDE 自动更新的链接（主题写法）
            doc.Header.Fields.Add(new FakeField(
                " DDEAUTO Excel \"[" + existingFile + "]总表\" \"R1C1:R3C3\" ", 10));

            // 浮动形状：OLE 链接对象
            doc.Shapes.Add(new FakeShape(existingFile, 320, 0));

            return new FakeApplication(doc);
        }

        private static LinkSourceInfo Find(List<LinkSourceInfo> list, string path)
        {
            for (int i = 0; i < list.Count; i++)
            {
                if (string.Equals(list[i].FullPath, path, StringComparison.OrdinalIgnoreCase))
                    return list[i];
            }
            return null;
        }

        private static void Check(ref int failed, string name, bool ok)
        {
            Console.WriteLine((ok ? "  [通过] " : "  [失败] ") + name);
            if (!ok) failed++;
        }
    }
}
