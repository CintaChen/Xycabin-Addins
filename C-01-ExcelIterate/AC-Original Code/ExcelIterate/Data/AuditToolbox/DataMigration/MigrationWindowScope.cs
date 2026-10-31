using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    // Never hide the Application or an existing user workbook. WPS may activate
    // its main frame during Open/Close even with ScreenUpdating disabled.
    internal sealed class MigrationWindowScope : IDisposable
    {
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr window);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr window, int command);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll")] private static extern bool LockSetForegroundWindow(uint code);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        private readonly IntPtr foreground;
        private readonly HashSet<uint> hostProcesses = new HashSet<uint>();
        internal string HostIdentity { get { return string.Join(",", hostProcesses); } }
        private readonly bool minimized;
        private readonly bool foregroundLock;
        [ThreadStatic] private static int lockDepth;
        internal MigrationWindowScope(object app)
        {
            foreground = GetForegroundWindow(); minimized = foreground != IntPtr.Zero && IsIconic(foreground);
            // Prevent Open/Save from taking foreground in the first place.
            // A user Alt+Tab/click automatically releases this Windows lock;
            // never acquire it for an unrelated foreground application.
            if (System.Windows.Forms.Form.FromHandle(foreground) is DataMigrationForm && (lockDepth > 0 || LockSetForegroundWindow(1)))
            { foregroundLock = true; lockDepth++; }
            try
            {
                IntPtr window = new IntPtr(Convert.ToInt64(((dynamic)app).Hwnd));
                AddHostWindow(window);
                using (var references = new ComBag())
                { dynamic active = references.Own(((dynamic)app).ActiveWindow); if (active != null) AddHostWindow(new IntPtr(Convert.ToInt64(active.Hwnd))); }
            }
            catch { } // Managed fixtures/non-window hosts do not need focus restoration.
        }
        private void AddHostWindow(IntPtr window)
        {
            foreach (IntPtr value in new[] { window, GetAncestor(window, 2), GetAncestor(window, 3) })
            { uint process; GetWindowThreadProcessId(value, out process); if (process != 0) hostProcesses.Add(process); }
        }
        public void Dispose()
        {
            if (foregroundLock && --lockDepth == 0) LockSetForegroundWindow(2);
            if (foreground == IntPtr.Zero || hostProcesses.Count == 0 || !IsWindow(foreground)) return;
            IntPtr current = GetForegroundWindow(); uint process; GetWindowThreadProcessId(current, out process);
            // Do not raise the dialog if the user switched to another program.
            if (current != IntPtr.Zero && current != foreground && !hostProcesses.Contains(process)) return;
            if (!minimized && IsIconic(foreground)) ShowWindow(foreground, 4); // SW_SHOWNOACTIVATE
            if (current != foreground) SetForegroundWindow(foreground);
        }
        internal static void SetOwnedVisible(object book, bool visible)
        {
            using (var references = new ComBag())
            {
                dynamic windows = references.Own(((dynamic)book).Windows);
                int count = Convert.ToInt32(windows.Count);
                if (count < 1) throw new InvalidOperationException("临时工作簿没有可管理的窗口。");
                for (int i = 1; i <= count; i++)
                {
                    dynamic window = references.Own(windows.Item(i)); window.Visible = visible;
                    if (ComFormulaHost.Boolean(window.Visible, "Visible") != visible)
                        throw new InvalidOperationException("当前 WPS 未能隐藏临时工作簿窗口，已停止本目标。");
                }
            }
        }
        internal static void RequireVisibleHost(object app)
        {
            if (!Marshal.IsComObject(app)) return;
            using (var references = new ComBag())
            {
                dynamic books = references.Own(((dynamic)app).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++)
                {
                    dynamic book = references.Own(books.Item(i)), windows = references.Own(book.Windows);
                    for (int j = 1; j <= Convert.ToInt32(windows.Count); j++)
                    { dynamic window = references.Own(windows.Item(j)); if (ComFormulaHost.Boolean(window.Visible, "Visible")) return; }
                }
            }
            // Older WPS may refuse automation calls after closing its last
            // visible view while hidden workbooks still exist. Fail before opening
            // any copies, instead of leaving an unresponsive invisible session.
            throw new InvalidOperationException("请先在 WPS 中打开并保留一个可见工作簿（例如源工作簿），再进行预检；请关闭需要更新的目标工作簿。");
        }
    }
}
