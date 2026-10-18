using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WordUpdater
{
    // LINK/DDE may start a hidden ET server, but a global PID difference does
    // not establish ownership. This compatibility entry point diagnoses only;
    // none of its callers explicitly created or owns the candidate process.
    internal static class EtCleanup
    {
        private delegate bool EnumWindowsProc(IntPtr window, IntPtr parameter);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr window);

        public static HashSet<int> SnapshotEtPids()
        {
            Process[] processes;
            try { processes = Process.GetProcessesByName("et"); }
            catch (Exception error) { Connect.Log("ET快照失败，保留所有实例：" + error.Message); return null; }
            var result = new HashSet<int>();
            bool complete = true;
            foreach (Process process in processes)
            {
                try { result.Add(process.Id); }
                catch (Exception error) { complete = false; Connect.Log("ET快照不完整：" + error.Message); }
                finally { process.Dispose(); }
            }
            return complete ? result : null;
        }

        public static int CleanupSpawnedHiddenEt(HashSet<int> before)
        {
            if (before == null) { Connect.Log("ET快照不可用；未关闭任何实例。"); return 0; }
            Process[] processes;
            try { processes = Process.GetProcessesByName("et"); }
            catch (Exception error) { Connect.Log("ET检查失败：" + error.Message); return 0; }
            var candidates = new List<int>();
            foreach (Process process in processes)
            {
                try { if (!process.HasExited) candidates.Add(process.Id); }
                catch (Exception error) { Connect.Log("ET状态读取失败，保留实例：" + error.Message); }
                finally { process.Dispose(); }
            }
            return ReviewCandidates(before, candidates, HasVisibleWindow, Connect.Log);
        }

        // Injected candidates test concurrency and incomplete enumeration
        // without attaching to, closing, or killing an actual application.
        private static int ReviewCandidates(HashSet<int> before, IEnumerable<int> current,
                                            Func<int, bool?> visible, Action<string> report)
        {
            if (before == null) { report("ET快照不可用；未关闭任何实例。"); return 0; }
            var seen = new HashSet<int>();
            foreach (int pid in current)
            {
                if (before.Contains(pid) || !seen.Add(pid)) continue;
                bool? state;
                try { state = visible(pid); }
                catch (Exception error) { report("ET窗口检查失败，保留pid=" + pid + "：" + error.Message); continue; }
                if (state == true) continue;
                report((state.HasValue ? "发现新增隐藏ET" : "ET窗口状态无法确认") +
                       "，pid=" + pid + "；无法证明由本次操作独占创建，保留实例，未发送关闭或终止请求。");
            }
            return 0;
        }

        private static bool? HasVisibleWindow(int pid)
        {
            bool visible = false, unknown = false;
            try
            {
                bool completed = EnumWindows(delegate(IntPtr window, IntPtr parameter)
                {
                    uint owner;
                    if (GetWindowThreadProcessId(window, out owner) == 0) unknown = true;
                    if (owner == (uint)pid && IsWindowVisible(window)) { visible = true; return false; }
                    return true;
                }, IntPtr.Zero);
                if (visible) return true;
                return completed && !unknown ? (bool?)false : null;
            }
            catch { return null; }
        }
    }
}
