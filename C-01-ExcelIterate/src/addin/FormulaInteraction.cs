using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // This registry and every COM reference in it belong to the Ribbon host STA.
    // The editor receives only a numeric ticket, an operation and an HWND.
    internal static class FormulaInteraction
    {
        private sealed class Session
        {
            internal long Id;
            internal IFormulaHost Host;
            internal IntPtr Owner;
            internal FormulaWindowThread Window;
            // Host STA only, never handed to the editor thread.
            internal object Application;
        }
        private static Control dispatcher;
        private static int hostThread;
        private static long sequence;
        private static Session current;
        internal static bool HasPending { get { return current != null; } }
        internal static Action<FormulaCommandResult> ResultForProbe;
        internal static string RecoveryForProbe;
        // Supplies the range for the custom flow without a real dialog, so probes can
        // walk both the accepted and the cancelled path.
        internal static Func<object> RangePickerForProbe;

        private static void EnsureDispatcher()
        {
            if (dispatcher != null) { AssertHostThread(); return; }
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("公式命令必须从宿主界面线程打开。");
            // COM hosts need not supply a WinForms SynchronizationContext. Do not
            // leave the automatically installed context behind in the host.
            SynchronizationContext previous = SynchronizationContext.Current;
            try { dispatcher = new Control(); var handle = dispatcher.Handle; hostThread = Thread.CurrentThread.ManagedThreadId; }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
        private static void AssertHostThread()
        {
            if (hostThread != Thread.CurrentThread.ManagedThreadId)
                throw new InvalidOperationException("公式目标只能在原宿主线程访问。");
        }
        internal static bool ActivatePending()
        {
            if (current == null) return false;
            AssertHostThread(); current.Window.Activate(); return true;
        }
        // Takes ownership even if capture or editor startup fails.
        internal static long Open(IFormulaHost host, FormulaOperation operation, IntPtr owner, object application = null)
        {
            bool transferred = false;
            try
            {
                EnsureDispatcher();
                if (ActivatePending()) return current.Id;
                // The custom flow picks its range after the template is accepted, so
                // nothing is captured up front and the sheet stays untouched meanwhile.
                if (operation != FormulaOperation.Custom) host.Collect();
                var session = new Session { Id = ++sequence, Host = host, Owner = owner, Application = application };
                session.Window = new FormulaWindowThread(session.Id, operation, owner);
                current = session; transferred = true;
                FormulaWindowState.Record("formula-editor-open", owner);
                session.Window.Start();
                return session.Id;
            }
            catch
            {
                if (transferred) { current = null; transferred = false; }
                throw;
            }
            finally { if (!transferred) host.Dispose(); }
        }
        // Called by an editor STA. No host object or host-capturing delegate crosses
        // that boundary; the ticket is resolved only after reaching the host STA.
        internal static void Receive(long ticket, FormulaRequest request)
        {
            Control target = dispatcher;
            if (target == null || target.IsDisposed) return;
            try { target.BeginInvoke((MethodInvoker)delegate { Complete(ticket, request); }); }
            catch (InvalidOperationException error) { UiDiagnostics.Record("formula-delivery", error); }
        }
        private static void Complete(long ticket, FormulaRequest request)
        {
            AssertHostThread();
            Session session = current;
            if (session == null || session.Id != ticket) return;
            // Remove before processing, so duplicate/late notifications cannot write.
            current = null;
            FormulaWindowState.Record(request == null ? "formula-cancel" : "formula-submit", session.Owner);
            if (request == null) { Release(session); return; }
            // The custom editor hides itself before this point, so the sheet is visible
            // while the user picks the range to write into. A cancelled pick must return
            // before the try below: the session goes back and the window comes back.
            if (request.Operation == FormulaOperation.Custom)
            {
                object range = PickRange(session.Application);
                if (range == null)
                {
                    current = session;
                    FormulaWindowState.Record("formula-pick-cancel", session.Owner);
                    session.Window.Show();
                    return;
                }
                IFormulaHost stale = session.Host; session.Host = null;
                if (stale != null) try { stale.Dispose(); } catch (Exception error) { UiDiagnostics.Record("formula-release", error); }
                session.Host = new ComFormulaHost(session.Application, FormulaCapture.Values, range);
            }
            if (!Connect.TryBeginFormulaExecution())
            {
                Release(session); session.Window.Close();
                FormulaWindowThread.Notice("其他操作正在执行，本次没有修改单元格。请稍后重新打开自定义参数。", session.Owner, true);
                return;
            }
            try
            {
                // Engine owns disposal once it starts, including failures and rollback.
                IFormulaHost host = session.Host; session.Host = null;
                FormulaCommandResult result = FormulaCommandEngine.Run(host, request, Connect.Log,
                    RecoveryForProbe ?? System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Excel_Iterate", "FormulaRecovery"));
                if (ResultForProbe != null) ResultForProbe(result); else Connect.PresentFormulaResult(result, session.Owner);
            }
            catch (Exception error)
            {
                UiDiagnostics.Record("formula-apply", error);
                FormulaWindowThread.Notice("一键公式发生错误：" + error.Message, session.Owner, true);
            }
            finally { Connect.EndFormulaExecution(); Release(session); session.Window.Close(); FormulaWindowState.Record("formula-finished", session.Owner); }
        }
        // Runs the native range picker on the host STA. Returns a Range for the caller to
        // own, or null when the user cancels or picks something that is not a range.
        private static object PickRange(object application)
        {
            if (RangePickerForProbe != null) return RangePickerForProbe();
            if (application == null) return null;
            object result = null;
            try
            {
                // Type 8 asks for a Range reference. A cancelled dialog yields the boolean
                // false instead of an object, so both shapes have to be handled here.
                result = ((dynamic)application).InputBox("请框选要应用公式的单元格区域，确定后开始写入。",
                    "自定义函数 · 选择区域", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            }
            catch (Exception error)
            {
                UiDiagnostics.Record("formula-range-pick", error);
                return null;
            }
            if (result == null || result is bool) return null;
            if (!LooksLikeRange(result)) { ReleaseCom(result); return null; }
            return result;
        }
        // The picker can hand back a shape or a chart; only a Range exposes Cells.
        private static bool LooksLikeRange(object candidate)
        {
            object cells = null;
            try { cells = ((dynamic)candidate).Cells; return cells != null; }
            catch (Exception) { return false; }
            finally { ReleaseCom(cells); }
        }
        private static void ReleaseCom(object value)
        { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }

        private static void Release(Session session)
        {
            IFormulaHost host = session.Host; session.Host = null;
            if (host != null) try { host.Dispose(); } catch (Exception error) { UiDiagnostics.Record("formula-release", error); }
        }
        internal static void Shutdown()
        {
            if (dispatcher == null) { FormulaWindowThread.CloseAll(); return; }
            AssertHostThread();
            Session session = current; current = null;
            FormulaWindowThread.CloseAll();
            if (session != null) Release(session);
            Control old = dispatcher; dispatcher = null; old.Dispose(); hostThread = 0;
        }
    }

    internal static class FormulaWindowState
    {
        [StructLayout(LayoutKind.Sequential)] private struct GuiInfo
        {
            internal int Size; internal uint Flags;
            internal IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret;
            internal int Left, Top, Right, Bottom;
        }
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr handle);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr handle, uint flags);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint process);
        [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint thread, ref GuiInfo info);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        internal static IntPtr ValidateOwner(IntPtr handle)
        {
            if (handle == IntPtr.Zero || !IsWindow(handle)) return IntPtr.Zero;
            IntPtr root = GetAncestor(handle, 2); uint process;
            uint thread = GetWindowThreadProcessId(root, out process);
            return root != IntPtr.Zero && process == (uint)Process.GetCurrentProcess().Id &&
                thread == GetCurrentThreadId() && IsWindowVisible(root) ? root : IntPtr.Zero;
        }
        internal static IntPtr ResolveOwner(object application)
        {
            object window = null;
            try
            {
                window = ((dynamic)application).ActiveWindow;
                IntPtr valid = ValidateOwner(new IntPtr(Convert.ToInt64(((dynamic)window).Hwnd)));
                if (valid != IntPtr.Zero) return valid;
            }
            catch { }
            finally { if (window != null && Marshal.IsComObject(window)) Marshal.ReleaseComObject(window); }
            IntPtr result = ValidateOwner(Process.GetCurrentProcess().MainWindowHandle);
            return result != IntPtr.Zero ? result : ValidateOwner(GetForegroundWindow());
        }
        internal static void Record(string stage, IntPtr owner)
        {
            try
            {
                uint thread = GetCurrentThreadId(); var info = new GuiInfo { Size = Marshal.SizeOf(typeof(GuiInfo)) };
                bool read = GetGUIThreadInfo(thread, ref info);
                UiDiagnostics.RecordState(stage, thread, owner, owner != IntPtr.Zero && IsWindowEnabled(owner),
                    info.Active, info.Focus, info.Capture, read ? info.Flags : uint.MaxValue);
            }
            catch (Exception error) { UiDiagnostics.Record("formula-window-state", error); }
        }
    }
}
