using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WordUpdater.Tools
{
    // =====================================================================
    // LinkViewFormProbe.exe —— 「查看链接」窗口填充逻辑的离线自检
    //
    //   守护的 bug（v1.3.1 修复，v1.3.2 加本探针）：
    //     只有一个源文件时，下方明细永远空白 —— 第一行被 DataGridView
    //     自动选中发生在 Tag 赋值之前，那次 SelectionChanged 里明细读
    //     到 null；之后点击已选中的行不再触发事件，FillItems 永远停
    //     在空状态。修复 = 填充完显式补一次 FillItems()。
    //
    //   手法：LinkViewForm 是 internal，用反射实例化（不 Show()），
    //   喂假对象模型，手动触发 OnLoad，再反射读私有 gridItem 行数。
    // =====================================================================

    public class LinkViewFormProbe
    {
        private static int checks;
        // ----------------------- 最小假对象模型（public，dynamic 需要） -----------------------

        public class FCode
        {
            public string Text;
            public int Start;
            public int HighlightColorIndex;
            public void Select() { }
            public FCode(string t, int s) { Text = t; Start = s; }
        }

        public class FField
        {
            public int Type = 56;
            public FCode Code;
            public FCode Result;
            public bool ShowCodes;
            public FField(string c, int s) { Code = new FCode(c, s); Result = new FCode("cached result", s + 1); }
            public object Update() { return true; }
            public void Unlink() { }
        }

        public class FColl
        {
            private readonly List<object> _l = new List<object>();
            public int Count { get { return _l.Count; } }
            public void Add(object o) { _l.Add(o); }
            public object this[int i] { get { return _l[i - 1]; } }
        }

        public class FRange
        {
            public FColl Fields = new FColl();
            public FColl InlineShapes = new FColl();
            public object NextStoryRange { get { return null; } }
        }

        public class FDoc
        {
            public string Name = "窗体自检.docx";
            public FColl StoryRanges = new FColl();
            public FColl Shapes = new FColl();
            public FRange Main = new FRange();
            public FDoc() { StoryRanges.Add(Main); }
        }

        public class FApp
        {
            public FDoc ActiveDocument;
            public object DisplayAlerts;
            public object ScreenUpdating;
            public FApp(FDoc d) { ActiveDocument = d; }
            public void Activate() { }
        }

        // ------------------------------- 主流程 -------------------------------

        // rule=R-25
        private static void CheckDialogStyleRedLines(ref int failed)
        {
            try
            {
                string dir = AppDomain.CurrentDomain.BaseDirectory;
                string addin = null;
                for (int i = 0; i < 6 && dir != null; i++)
                {
                    string candidate = Path.Combine(dir, "src", "addin");
                    if (Directory.Exists(candidate)) { addin = candidate; break; }
                    dir = Path.GetDirectoryName(dir);
                }
                if (addin == null) { Console.WriteLine("[跳过] 找不到 src/addin，弹窗样式红线未扫描"); return; }
                string file = Path.Combine(addin, "LinkViewForm.cs");
                Check(ref failed, "LinkViewForm.cs 存在", WpsShared.SourcePaths.Exists(file));
                if (!WpsShared.SourcePaths.Exists(file)) return;
                var code = new StringBuilder();
                foreach (string line in File.ReadAllLines(file))
                {
                    string t = line.TrimStart();
                    if (t.StartsWith("//") || t.StartsWith("*") || t.StartsWith("/*")) continue;
                    code.AppendLine(line);
                }
                string text = code.ToString();
                Check(ref failed, "LinkViewForm 无 TopMost=true 赋值",
                    !Regex.IsMatch(text, @"^\s*(?:this\.)?TopMost\s*=\s*true\s*;", RegexOptions.Multiline));
                Check(ref failed, "LinkViewForm 无工具窗样式赋值",
                    !Regex.IsMatch(text, @"^\s*(?:this\.)?(ShowInTaskbar|MinimizeBox)\s*=\s*false\s*;", RegexOptions.Multiline));
            }
            catch (Exception ex)
            {
                Check(ref failed, "弹窗样式红线扫描可执行", false);
                Console.WriteLine("[样式扫描异常] " + ex.Message);
            }
        }

        [STAThread]
        public static int Main(string[] args)
        {
            try { Console.OutputEncoding = Encoding.UTF8; }
            catch { }
            int failed = 0;
            CheckDialogStyleRedLines(ref failed);

            string token = Guid.NewGuid().ToString("N");
            string f1 = Path.Combine(Path.GetTempPath(), "FormProbe_" + token + "_源一.xlsx");
            string f2 = Path.Combine(Path.GetTempPath(), "FormProbe_" + token + "_源二.xlsx");
            try
            {
                File.WriteAllText(f1, "xlsx");
                File.WriteAllText(f2, "xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("[准备失败] 无法创建临时文件：" + ex.Message);
                return 2;
            }

            Type formType = typeof(LinkViewEngine).Assembly.GetType("WordUpdater.LinkViewForm");
            Check(ref failed, "应能从程序集反射到 LinkViewForm 类型", formType != null);
            if (formType == null) return 1;

            Console.WriteLine("================ 场景 A：只有一个源文件 ================");
            // 回归目标：修复前此时 gridItem 为 0 行
            FDoc docA = new FDoc();
            docA.Main.Fields.Add(new FField(
                " LINK Excel.Sheet.12 \"" + f1 + "\" \"Sheet1!R1C1\" \\t ", 100));
            Form formA = CreateForm(formType, new FApp(docA));
            DataGridView gridSrcA = GetGrid(formType, formA, "gridSrc");
            DataGridView gridItemA = GetGrid(formType, formA, "gridItem");
            FireLoad(formA);
            Console.WriteLine(string.Format("源文件行数={0}  明细行数={1}",
                gridSrcA.Rows.Count, gridItemA.Rows.Count));
            Check(ref failed, "单源文件时明细应立即显示 1 行（修复前为 0）",
                gridItemA.Rows.Count == 1);
            Check(ref failed, "首行应处于选中状态", gridSrcA.SelectedRows.Count == 1);

            Console.WriteLine();
            Console.WriteLine("================ 场景 B：两个源文件，切换选中行 ================");
            FDoc docB = new FDoc();
            docB.Main.Fields.Add(new FField(
                " LINK Excel.Sheet.12 \"" + f1 + "\" \"Sheet1!R1C1\" \\t ", 100));
            docB.Main.Fields.Add(new FField(
                " LINK Excel.Sheet.12 \"" + f2 + "\" \"Sheet2!R1C1\" \\t ", 160));
            docB.Main.Fields.Add(new FField(
                " LINK Excel.Sheet.12 \"" + f2 + "\" \"Sheet2!R2C2\" \\t ", 200));
            Form formB = CreateForm(formType, new FApp(docB));
            DataGridView gridSrcB = GetGrid(formType, formB, "gridSrc");
            DataGridView gridItemB = GetGrid(formType, formB, "gridItem");
            FireLoad(formB);
            Console.WriteLine(string.Format("初始：源行数={0}  明细行数={1}",
                gridSrcB.Rows.Count, gridItemB.Rows.Count));
            Check(ref failed, "打开后明细应立即显示第一源的 1 条", gridItemB.Rows.Count == 1);

            gridSrcB.Rows[1].Selected = true;   // 模拟用户点第二行
            Console.WriteLine("切到第二行后：明细行数=" + gridItemB.Rows.Count);
            Check(ref failed, "切到第二源后明细应显示它的 2 条", gridItemB.Rows.Count == 2);

            gridSrcB.Rows[0].Selected = true;   // 点回第一行
            Check(ref failed, "点回第一源后明细应恢复 1 条", gridItemB.Rows.Count == 1);

            Console.WriteLine();
            Console.WriteLine("================ 场景 C：连续定位先还原旧高亮 ================");
            RepeatedJumpChecks(formType, f1, ref failed);

            formA.Dispose(); formB.Dispose();

            try { File.Delete(f1); } catch { }
            try { File.Delete(f2); } catch { }

            Console.WriteLine();
            Console.WriteLine(failed == 0 ? "全部自检通过：" + checks + "项。"
                : string.Format("有 {0} 项自检未通过。", failed));
            return failed == 0 ? 0 : 1;
        }

        // ------------------------------- 反射辅助 -------------------------------

        private static void RepeatedJumpChecks(Type formType, string source, ref int failed)
        {
            FDoc doc = new FDoc();
            FField first = new FField(" LINK Excel.Sheet.12 \"" + source + "\" \"Sheet1!R1C1\" \\t ", 100);
            FField second = new FField(" LINK Excel.Sheet.12 \"" + source + "\" \"Sheet1!R2C1\" \\t ", 200);
            doc.Main.Fields.Add(first); doc.Main.Fields.Add(second);
            using (Form form = CreateForm(formType, new FApp(doc)))
            {
                try
                {
                    FireLoad(form);
                    DataGridView grid = GetGrid(formType, form, "gridItem");
                    MethodInfo jump = formType.GetMethod("JumpToSelectedItem", BindingFlags.Instance | BindingFlags.NonPublic);
                    Check(ref failed, "两字段同一来源，回归调用生产JumpToSelectedItem", grid.Rows.Count == 2 && jump != null);
                    if (grid.Rows.Count != 2 || jump == null) return;
                    foreach (bool codes in new[] { false, true }) foreach (int original in new[] { 0, 2, 7 })
                    {
                        ClearPendingFlash(); first.ShowCodes = codes;
                        FCode target = codes ? first.Code : first.Result; target.HighlightColorIndex = original;
                        grid.ClearSelection(); grid.Rows[0].Selected = true;
                        for (int repeat = 0; repeat < 3; repeat++)
                        {
                            jump.Invoke(form, null);
                            Check(ref failed, "同字段第" + (repeat + 1) + "次定位仍标黄；原色=" + original + "，ShowCodes=" + codes, target.HighlightColorIndex == 7);
                            Check(ref failed, "每次只登记当前一条高亮", PendingFlashCount() == 1);
                        }
                        ClearPendingFlash();
                        Check(ref failed, "最终按原色恢复而非临时黄或清零；原色=" + original + "，ShowCodes=" + codes, target.HighlightColorIndex == original);
                    }
                    first.ShowCodes = second.ShowCodes = false;
                    first.Result.HighlightColorIndex = 2; second.Result.HighlightColorIndex = 5;
                    grid.ClearSelection(); grid.Rows[0].Selected = true; jump.Invoke(form, null);
                    grid.ClearSelection(); grid.Rows[1].Selected = true; jump.Invoke(form, null);
                    Check(ref failed, "换字段先恢复旧原色，再保持新字段黄色", first.Result.HighlightColorIndex == 2 && second.Result.HighlightColorIndex == 7);
                    ClearPendingFlash();
                    Check(ref failed, "换字段后第二字段恢复自身非零原色", second.Result.HighlightColorIndex == 5);
                    first.Result.HighlightColorIndex = 2; grid.ClearSelection(); grid.Rows[0].Selected = true; jump.Invoke(form, null);
                    LinkItemInfo stale = (LinkItemInfo)grid.Rows[0].Tag;
                    string raw = stale.RawCode; stale.RawCode = "expired field code";
                    // The production form would show a warning for this path. Core
                    // refusal is tested without invoking that warning dialog.
                    ClearPendingFlash(); object range; int previous;
                    string error = LinkViewEngine.JumpToItemEx(new FApp(doc), stale, out range, out previous);
                    stale.RawCode = raw;
                    Check(ref failed, "旧明细拒绝定位时已恢复前一高亮且不生成新高亮", error != null && range == null && first.Result.HighlightColorIndex == 2);
                }
                finally { ClearPendingFlash(); }
            }
        }

        private static Type FlashOwner()
        { return typeof(LinkViewEngine).Assembly.GetType("WordUpdater.Connect", true); }
        private static void ClearPendingFlash()
        { FlashOwner().GetMethod("ClearPendingFlash", BindingFlags.Public | BindingFlags.Static).Invoke(null, null); }
        private static int PendingFlashCount()
        {
            object session = FlashOwner().GetProperty("SharedFlash", BindingFlags.Public | BindingFlags.Static).GetValue(null, null);
            return ((FlashSession)session).Count;
        }

        private static Form CreateForm(Type formType, object app)
        {
            return (Form)Activator.CreateInstance(formType,
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                null, new object[] { app }, null);
        }

        /// <summary>手动触发 Load 事件（Form 未 Show 时不会自动触发）。</summary>
        private static void FireLoad(Form form)
        {
            MethodInfo m = typeof(Form).GetMethod("OnLoad",
                BindingFlags.NonPublic | BindingFlags.Instance);
            m.Invoke(form, new object[] { EventArgs.Empty });
        }

        private static DataGridView GetGrid(Type formType, Form form, string name)
        {
            FieldInfo f = formType.GetField(name,
                BindingFlags.NonPublic | BindingFlags.Instance);
            if (f == null) throw new InvalidOperationException("字段不存在：" + name);
            return (DataGridView)f.GetValue(form);
        }

        private static void Check(ref int failed, string name, bool ok)
        {
            checks++;
            Console.WriteLine((ok ? "  [通过] " : "  [失败] ") + name);
            if (!ok) failed++;
        }
    }
}
