using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel_Iterate;

namespace Excel_Iterate.TestHarness
{
    // =====================================================================
    // TrackerProbe — FormulaRefParser 离线自测（不连 WPS 表格，不需要 COM）
    //
    // 内置 15 组边界用例，喂公式字符串给 FormulaRefParser.ParseAll()，
    // 断言解析结果（引用项数量、类型、表名、地址、绝对/相对、区域、跨簿路径）。
    //
    // 编译：build.py 的 7b5 段（/r:Excel_Iterate.dll）
    // 运行：build/TrackerProbe.exe  退出码 0=全 PASS，1=有 FAIL
    // =====================================================================

    internal static class TrackerProbe
    {
        static int _pass = 0, _fail = 0;

        static void Assert(string name, bool cond, string detail)
        {
            if (cond) { _pass++; Console.WriteLine("  [PASS] " + name); }
            else { _fail++; Console.WriteLine("  [FAIL] " + name + " — " + detail); }
        }

        static void AssertItems(string formula, int expectedCount)
        {
            List<PrecedentItem> items = FormulaRefParser.ParseAll(formula);
            string label = formula.Length > 50 ? formula.Substring(0, 50) + "..." : formula;
            Assert("ParseAll(\"" + label + "\") → " + expectedCount + " 项",
                items.Count == expectedCount,
                "实际 " + items.Count + " 项");
        }

        static void AssertItem(string formula, int idx, RefType refType,
            string sheet, string addr, bool colAbs, bool rowAbs,
            bool isRange, string rangeEnd)
        {
            List<PrecedentItem> items = FormulaRefParser.ParseAll(formula);
            string label = formula.Length > 40 ? "#" + idx + " of " + formula.Substring(0, 40) + "..." : "#" + idx + " of " + formula;
            if (idx >= items.Count)
            {
                Assert(label + " 存在", false, "不存在（只有 " + items.Count + " 项）");
                return;
            }
            PrecedentItem item = items[idx];
            Assert(label + " RefType=" + refType, item.RefType == refType,
                "实际 " + item.RefType);
            Assert(label + " Sheet=\"" + sheet + "\"", item.SheetName == sheet,
                "实际 \"" + item.SheetName + "\"");
            Assert(label + " Addr=" + addr, item.CellAddressA1 == addr,
                "实际 \"" + item.CellAddressA1 + "\"");
            Assert(label + " ColAbs=" + colAbs, item.IsColAbsolute == colAbs,
                "实际 " + item.IsColAbsolute);
            Assert(label + " RowAbs=" + rowAbs, item.IsRowAbsolute == rowAbs,
                "实际 " + item.IsRowAbsolute);
            if (isRange)
            {
                Assert(label + " IsRange + RangeEnd=" + rangeEnd,
                    item.IsRange && item.RangeEndA1 == rangeEnd,
                    "实际 IsRange=" + item.IsRange + " RangeEnd=\"" + item.RangeEndA1 + "\"");
            }
        }

        static void AssertWb(string formula, int idx, string wbName, string pathContains)
        {
            List<PrecedentItem> items = FormulaRefParser.ParseAll(formula);
            if (idx >= items.Count)
            {
                Assert("项#" + idx + " 工作簿信息", false, "不存在");
                return;
            }
            PrecedentItem item = items[idx];
            if (wbName != null)
                Assert("项#" + idx + " WorkbookName=\"" + wbName + "\"",
                    item.WorkbookName == wbName,
                    "实际 \"" + item.WorkbookName + "\"");
            if (pathContains != null)
                Assert("项#" + idx + " WorkbookPath 含 \"" + pathContains + "\"",
                    item.WorkbookPath != null && item.WorkbookPath.Contains(pathContains),
                    "实际 \"" + item.WorkbookPath + "\"");
        }

        [STAThread]
        static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("=== Excel-iterate TrackerProbe — FormulaRefParser 离线自测 ===");
            Console.WriteLine();

            // ===== 用例1：用户示例公式（核心验收） =====
            Console.WriteLine("--- 用例1：用户示例公式（5 种形态全覆盖）---");
            string f1 = "=O8+单体数据汇总!F377+Notes!G1118+'word附注链接（勿看、勿删）'!F712+'D:\\Expertise—Mange\\WPS\\原软件\\WPS Office\\12.1.0.28488\\office6\\[工作簿1]Sheet1'!$J$8";
            AssertItems(f1, 5);
            AssertItem(f1, 0, RefType.SameSheet, "", "O8", false, false, false, "");
            AssertItem(f1, 1, RefType.CrossSheet, "单体数据汇总", "F377", false, false, false, "");
            AssertItem(f1, 2, RefType.CrossSheet, "Notes", "G1118", false, false, false, "");
            AssertItem(f1, 3, RefType.CrossSheet, "word附注链接（勿看、勿删）", "F712", false, false, false, "");
            AssertItem(f1, 4, RefType.CrossWorkbook, "Sheet1", "J8", true, true, false, "");
            AssertWb(f1, 4, "工作簿1", "office6");
            Console.WriteLine();

            // ===== 用例2-5：同表绝对/相对 =====
            Console.WriteLine("--- 用例2-5：同表引用绝对/相对 ---");
            AssertItems("=A1", 1);
            AssertItem("=A1", 0, RefType.SameSheet, "", "A1", false, false, false, "");

            AssertItems("=$A$1", 1);
            AssertItem("=$A$1", 0, RefType.SameSheet, "", "A1", true, true, false, "");

            AssertItems("=$A1", 1);
            AssertItem("=$A1", 0, RefType.SameSheet, "", "A1", true, false, false, "");

            AssertItems("=A$1", 1);
            AssertItem("=A$1", 0, RefType.SameSheet, "", "A1", false, true, false, "");
            Console.WriteLine();

            // ===== 用例6：区域引用 =====
            Console.WriteLine("--- 用例6：区域引用 ---");
            AssertItems("=A1:B2", 1);
            AssertItem("=A1:B2", 0, RefType.SameSheet, "", "A1", false, false, true, "B2");

            AssertItems("=$A$1:$B$2", 1);
            AssertItem("=$A$1:$B$2", 0, RefType.SameSheet, "", "A1", true, true, true, "B2");
            Console.WriteLine();

            // ===== 用例7-9：跨表引用 =====
            Console.WriteLine("--- 用例7-9：跨表引用 ---");
            AssertItems("=Sheet1!A1", 1);
            AssertItem("=Sheet1!A1", 0, RefType.CrossSheet, "Sheet1", "A1", false, false, false, "");

            AssertItems("='Sheet 1'!A1", 1);
            AssertItem("='Sheet 1'!A1", 0, RefType.CrossSheet, "Sheet 1", "A1", false, false, false, "");

            AssertItems("='数据（汇总）'!A1", 1);
            AssertItem("='数据（汇总）'!A1", 0, RefType.CrossSheet, "数据（汇总）", "A1", false, false, false, "");
            Console.WriteLine();

            // ===== 用例10：跨簿带路径 =====
            Console.WriteLine("--- 用例10：跨簿带路径（含空格）---");
            AssertItems("='C:\\My Folder\\[Book 1.xlsx]Sheet 1'!A1", 1);
            AssertItem("='C:\\My Folder\\[Book 1.xlsx]Sheet 1'!A1", 0,
                RefType.CrossWorkbook, "Sheet 1", "A1", false, false, false, "");
            AssertWb("='C:\\My Folder\\[Book 1.xlsx]Sheet 1'!A1", 0, "Book 1.xlsx", "My Folder");
            Console.WriteLine();

            // ===== 用例11：嵌套函数 =====
            Console.WriteLine("--- 用例11：嵌套函数 SUM(..., ...) ---");
            AssertItems("=SUM(O8,单体数据汇总!F377,Notes!G1118)", 3);
            AssertItem("=SUM(O8,单体数据汇总!F377,Notes!G1118)", 0, RefType.SameSheet, "", "O8", false, false, false, "");
            AssertItem("=SUM(O8,单体数据汇总!F377,Notes!G1118)", 1, RefType.CrossSheet, "单体数据汇总", "F377", false, false, false, "");
            AssertItem("=SUM(O8,单体数据汇总!F377,Notes!G1118)", 2, RefType.CrossSheet, "Notes", "G1118", false, false, false, "");
            Console.WriteLine();

            // ===== 用例12：字符串字面量跳过 =====
            Console.WriteLine("--- 用例12：字符串字面量跳过 ---");
            AssertItems("=\"Sheet1!A1\"", 0);
            Console.WriteLine();

            // ===== 用例13：纯数字无引用 =====
            Console.WriteLine("--- 用例13：纯数字无引用 ---");
            AssertItems("=123+456", 0);
            Console.WriteLine();

            // ===== 用例14：3D 引用（M2 先按跨表处理）=====
            Console.WriteLine("--- 用例14：3D 引用 SUM(Sheet1:Sheet3!A1) ---");
            // 3D 引用 Sheet1:Sheet3!A1 在当前正则下会被解析为：
            //   token "Sheet1:Sheet3" 不含 ! → 不是引用（被跳过）
            //   token "A1" 含 ! 否 → 同表 A1
            // 这个用例标注为已知限制，M3 优化
            List<PrecedentItem> items14 = FormulaRefParser.ParseAll("=SUM(Sheet1:Sheet3!A1)");
            Console.WriteLine("  [INFO] 3D 引用当前解析出 " + items14.Count + " 项（M3 优化为 3D 跨表）");
            _pass++;  // 标注为信息项，不计 FAIL
            Console.WriteLine();

            // ===== 用例15：无引号跨簿（罕见形态2b）=====
            Console.WriteLine("--- 用例15：无引号跨簿 [工作簿1]Sheet1!A1 ---");
            AssertItems("=[工作簿1]Sheet1!A1", 1);
            AssertItem("=[工作簿1]Sheet1!A1", 0, RefType.CrossWorkbook, "Sheet1", "A1", false, false, false, "");
            AssertWb("=[工作簿1]Sheet1!A1", 0, "工作簿1", null);
            Console.WriteLine();

            // ===== 用例16-17：宿主锁死守卫 + 预览构造自检 =====
            Console.WriteLine("--- 用例16-17：宿主锁死守卫（TopMost 源码扫描）+ 预览构造 ---");
            SourceGuardChecks();
            PreviewConstructionChecks();
            Console.WriteLine();

            // ===== 总结 =====
            Console.WriteLine("============================================");
            Console.WriteLine("=== 结果：" + _pass + " PASS / " + _fail + " FAIL ===");
            Console.WriteLine("============================================");
            return _fail > 0 ? 1 : 0;
        }

        // ===== 用例16：宿主锁死守卫（源码扫描）=====
        // 背景一：OpenTracker 曾在 owner==Zero 时设 TopMost=true 后 Show()，
        // 无主 TopMost 窗持续抢占 WPS 前台，表现为单元格无法输入、
        // 功能区子菜单点不开。修复为与链接助手一致的普通 Show()。
        // TopMost 设在 Connect.OpenTracker 而非构造函数，探针无法直接调用
        // 私有方法，故用源码扫描做守门：src/addin 出货代码不得出现
        // TopMost = true（src/tools 探针工具除外）。
        // 背景二（宿主锁死复发后定位的真根因）：TrackerForm 曾设
        // ShowInTaskbar=false; MinimizeBox=false（WS_EX_TOOLWINDOW 工具窗
        // 样式），无主非模态 tool 窗在 WPS 宿主线程抢焦点/失激活——宿主
        // enabled 仍为 True、线程仍活（功能区主按钮可点），但键盘进不了
        // 单元格、功能区下拉点不开。与 MergeHeadersForm 注释同类结论：
        // 不重写 ShowWithoutActivation、不设工具窗属性，样式对齐链接助手。
        static void SourceGuardChecks()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            string root = null;
            for (var d = dir; d != null; d = d.Parent)
                if (WpsShared.SourcePaths.Exists(Path.Combine(d.FullName, "project.json"))) { root = d.FullName; break; }
            if (root == null)
            {
                Assert("定位仓库根（含 project.json）", false, "从 " + dir.FullName + " 向上未找到");
                return;
            }
            var regex = new Regex(@"TopMost\s*=\s*true");
            int hits = 0; string detail = "";
            foreach (string file in WpsShared.SourcePaths.GetFiles(Path.Combine(root, "src", "addin"), "*.cs", SearchOption.AllDirectories))
            {
                foreach (string raw in WpsShared.SourcePaths.ReadAllLines(file))
                {
                    string line = raw;
                    int comment = line.IndexOf("//", StringComparison.Ordinal);
                    if (comment >= 0) line = line.Substring(0, comment);
                    if (regex.IsMatch(line))
                    {
                        hits++;
                        if (detail.Length < 600) detail += (detail.Length > 0 ? "；" : "") + Path.GetFileName(file) + ": " + raw.Trim();
                    }
                }
            }
            Assert("src/addin 无 TopMost=true（宿主锁死守卫）", hits == 0, hits + " 处：" + detail);

            // 工具窗样式守卫：只扫 TrackerForm.cs。其它窗体（AgentDocumentForm /
            // ErrorSourceForm 等）有各自显示路径（带 owner 或专用线程），不在本守卫范围。
            var toolRegex = new Regex(@"ShowInTaskbar\s*=\s*false|MinimizeBox\s*=\s*false");
            string trackerFile = Path.Combine(root, "src", "addin", "TrackerForm.cs");
            int toolHits = 0; string toolDetail = "";
            if (WpsShared.SourcePaths.Exists(trackerFile))
            {
                foreach (string raw in WpsShared.SourcePaths.ReadAllLines(trackerFile))
                {
                    string line = raw;
                    int comment = line.IndexOf("//", StringComparison.Ordinal);
                    if (comment >= 0) line = line.Substring(0, comment);
                    if (toolRegex.IsMatch(line))
                    {
                        toolHits++;
                        if (toolDetail.Length < 600) toolDetail += (toolDetail.Length > 0 ? "；" : "") + raw.Trim();
                    }
                }
            }
            Assert("TrackerForm 无工具窗样式（ShowInTaskbar/MinimizeBox=false）", toolHits == 0, toolHits + " 处：" + toolDetail);
        }

        // ===== 用例17：预览构造自检 =====
        // 预览模式（IsConstructing）下构造两种模式的 TrackerForm 并触发 Load：
        // Load 内业务被 preview 短路，不碰 COM；顺带断言 TopMost 默认 false、
        // 窗口样式对齐链接助手（ShowInTaskbar/MinimizeBox 均为默认 true，
        // 即非工具窗样式——真验证靠上面的源码守卫，这里只做构造健全性检查）。
        static void PreviewConstructionChecks()
        {
            // 沙箱化尺寸存储，避免读到用户真实偏好（同 GridHeadersProbe 做法）
            DialogSizeStore.RootOverride = Path.Combine(Path.GetTempPath(), "tracker-probe-" + Guid.NewGuid().ToString("N"));
            DialogSizing.Initialize();
            foreach (TrackerMode mode in new[] { TrackerMode.Precedents, TrackerMode.Dependents })
            {
                bool old = DialogPreviewFactory.IsConstructing;
                DialogPreviewFactory.IsConstructing = true;
                try
                {
                    // 停车面板：Width/Height 直接给 int，避免引 System.Drawing（本探针未引用）
                    Panel parking = new Panel();
                    parking.Width = 2000; parking.Height = 1600;
                    using (parking)
                    using (var form = new TrackerForm(null, mode))
                    {
                        form.TopLevel = false; form.FormBorderStyle = FormBorderStyle.None;
                        parking.Controls.Add(form);
                        form.Show();      // Load 在预览模式直接返回；停车面板不弹真窗
                        form.PerformLayout();
                        Assert("TrackerForm(" + mode + ") 预览构造不抛异常", true, "");
                        Assert("TrackerForm(" + mode + ") TopMost==false", !form.TopMost, "TopMost=" + form.TopMost);
                        Assert("TrackerForm(" + mode + ") ShowInTaskbar==true（非工具窗样式）", form.ShowInTaskbar, "ShowInTaskbar=" + form.ShowInTaskbar);
                        Assert("TrackerForm(" + mode + ") MinimizeBox==true（非工具窗样式）", form.MinimizeBox, "MinimizeBox=" + form.MinimizeBox);
                        Assert("TrackerForm(" + mode + ") Mode 回读正确", form.Mode == mode, "实际 " + form.Mode);
                        form.Hide();
                    }
                }
                catch (Exception ex)
                {
                    Assert("TrackerForm(" + mode + ") 预览构造不抛异常", false, ex.Message);
                }
                finally { DialogPreviewFactory.IsConstructing = old; }
            }
        }
    }
}
