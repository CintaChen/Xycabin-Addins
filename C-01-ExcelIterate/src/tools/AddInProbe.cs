using System;
using System.Collections.Generic;
using Id = WpsShared.ProjectInfo;
using System.Runtime.InteropServices;
using System.Text;

namespace WpsShared.Tools
{
    // Attach to an already running host and inspect its current registration.
    // Exit codes: 0 = inspected Ribbon XML, 1 = observed missing/disconnected/empty
    // add-in, 2 = INCOMPLETE (host or diagnostic object could not be inspected).
    // Borrowed COM references are released once per acquisition; never Quit().
    internal static class AddInProbe
    {
        private const string OWN = Id.AddinName;

        public static int Main(string[] args)
        {
            object host = null;
            string used = null;
            List<string> attachErrors = new List<string>();
            foreach (string p in Id.ProgIds)
            {
                try
                {
                    host = Marshal.GetActiveObject(p);
                    used = p;
                    break;
                }
                catch (Exception ex) { attachErrors.Add(p + ": " + Error(ex)); }
            }
            if (host == null)
            {
                foreach (string error in attachErrors) Console.WriteLine("  " + error);
                return Incomplete("没有找到可读取的运行中" + Id.HostDisplayName
                    + "。请打开对应宿主再运行本工具。");
            }
            Console.WriteLine("已连接：" + used);
            return Inspect(host, ReleaseBorrowed);
        }

        // Also supports isolated managed fixtures without active-host attachment.
        internal static int Inspect(object host, Action<object> releaseReference)
        {
            if (releaseReference == null) throw new ArgumentNullException("releaseReference");
            List<object> borrowed = new List<object>();
            Keep(borrowed, host);
            int result = 2;
            bool releaseFailed = false;
            try { result = InspectCore(host, borrowed); }
            catch (Exception ex) { result = Incomplete("读取宿主或加载项失败：" + Error(ex)); }
            finally
            {
                // Duplicate identities remain: each property/Item access acquired
                // one reference, even when COM returns the same RCW.
                for (int i = borrowed.Count - 1; i >= 0; i--)
                {
                    try { releaseReference(borrowed[i]); }
                    catch (Exception ex)
                    {
                        releaseFailed = true;
                        Incomplete("释放借用引用失败：" + Error(ex));
                    }
                }
            }
            return releaseFailed ? 2 : result;
        }

        private static int InspectCore(object host, List<object> borrowed)
        {
            if (host == null) return Incomplete("宿主对象不可用。");
            dynamic addins = Keep(borrowed, ((dynamic)host).COMAddIns);
            if (addins == null) return Incomplete("宿主没有提供 COMAddIns 集合。");
            int n = addins.Count;
            if (n < 0) return Incomplete("COMAddIns.Count 无效。");
            Console.WriteLine("COM 加载项共 " + n + " 个：");

            int found = -1;
            bool unreadableItem = false;
            for (int i = 1; i <= n; i++)
            {
                try
                {
                    dynamic ai = Keep(borrowed, addins.Item(i));
                    string prog = "" + ai.ProgId;
                    string desc;
                    try { desc = "" + ai.Description; }
                    catch (Exception ex) { desc = "<Description 读取失败：" + Error(ex) + ">"; }
                    string connected;
                    try { connected = ((bool)ai.Connect).ToString(); }
                    catch (Exception ex) { connected = "<读取失败：" + Error(ex) + ">"; }
                    Console.WriteLine("  " + i + ". " + desc + "  [" + prog + "]  Connect=" + connected);
                    // Registration uses exactly Id.AddinName. Substring and
                    // Description matches can select an unrelated add-in.
                    if (string.Equals(prog, OWN, StringComparison.OrdinalIgnoreCase))
                    {
                        if (found > 0) return Incomplete("发现多个相同 ProgId 的本加载项，无法唯一选择。");
                        found = i;
                    }
                }
                catch (Exception ex)
                {
                    unreadableItem = true;
                    Console.WriteLine("  " + i + ". <读取失败：" + Error(ex) + ">");
                }
            }
            if (found <= 0)
            {
                if (unreadableItem) return Incomplete("部分加载项无法枚举，尚不能确认是否存在 " + OWN + "。");
                Console.WriteLine("FAIL: COMAddIns 中没有找到本加载项（" + OWN + "）。");
                return 1;
            }

            dynamic mine = Keep(borrowed, addins.Item(found));
            if (!string.Equals("" + mine.ProgId, OWN, StringComparison.OrdinalIgnoreCase))
                return Incomplete("枚举后加载项集合发生变化，所选 ProgId 不再匹配。");
            string guid;
            try { guid = "" + mine.Guid; }
            catch (Exception ex) { guid = "<读取失败：" + Error(ex) + ">"; }
            bool on;
            try { on = (bool)mine.Connect; }
            catch (Exception ex) { return Incomplete("本加载项 Connect 无法读取：" + Error(ex)); }
            Console.WriteLine("本加载项：guid=" + guid + "  Connect=" + on + "  progId=" + OWN);
            if (!on)
            {
                Console.WriteLine("FAIL: Connect=false：加载项当前未连接。");
                return 1;
            }

            string ribbonId;
            if (string.Equals(Id.Host, "Excel", StringComparison.OrdinalIgnoreCase))
                ribbonId = "Microsoft.Excel.Workbook";
            else if (string.Equals(Id.Host, "Word", StringComparison.OrdinalIgnoreCase))
                ribbonId = "Microsoft.Word.Document";
            else return Incomplete("尚未配置此宿主的 RibbonID：" + Id.Host);

            object ribbon;
            try { ribbon = Keep(borrowed, mine.Object); }
            catch (Exception ex)
            {
                return Incomplete("Connect=true，但宿主未提供可读取的诊断 Object：" + Error(ex));
            }
            if (ribbon == null)
                return Incomplete("Connect=true，但宿主没有公开诊断 Object；尚未验证功能区 XML。");
            string xml;
            try { xml = ((dynamic)ribbon).GetCustomUI(ribbonId); }
            catch (Exception ex)
            {
                return Incomplete("Connect=true，但无法通过诊断 Object 调用 GetCustomUI：" + Error(ex));
            }
            if (string.IsNullOrEmpty(xml))
            {
                Console.WriteLine("FAIL: GetCustomUI(" + ribbonId + ") 返回空 XML。");
                return 1;
            }
            Console.WriteLine("GetCustomUI(" + ribbonId + ") 返回 XML " + xml.Length + " 字符，前 12 行：");
            string[] lines = xml.Replace("\r\n", "\n").Split('\n');
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < Math.Min(lines.Length, 12); i++) sb.AppendLine("    " + lines[i]);
            Console.Write(sb.ToString());
            int buttons = 0;
            foreach (string line in lines)
                if (line.IndexOf("<button", StringComparison.Ordinal) >= 0) buttons++;
            Console.WriteLine("包含按钮的 XML 行数：" + buttons);
            if (buttons == 0) Console.WriteLine("FAIL: XML 中未找到按钮。");
            return buttons > 0 ? 0 : 1;
        }

        private static object Keep(List<object> borrowed, object value)
        {
            if (value != null) borrowed.Add(value);
            return value;
        }

        private static void ReleaseBorrowed(object value)
        {
            if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }

        private static string Error(Exception ex)
        {
            return ex.GetType().Name + ": " + ex.Message;
        }

        private static int Incomplete(string message)
        {
            Console.WriteLine("INCOMPLETE: " + message);
            return 2;
        }
    }
}
