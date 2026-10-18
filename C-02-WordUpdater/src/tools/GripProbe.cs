using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

// Real controls and capture transitions on a private desktop; never Show or SwitchDesktop.
internal static class GripProbe
{
    const BindingFlags All = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
    static int checks;
    static void Check(bool value, string name) { if (!value) throw new Exception(name); checks++; }
    static object Make(Type type, params object[] args) { return Activator.CreateInstance(type, All, null, args, null); }
    static void Event(Control control, string method, params object[] args)
    {
        try { control.GetType().GetMethod(method, All).Invoke(control, args); }
        catch (TargetInvocationException e) { throw e.InnerException; }
    }
    static void Mouse(Control grip, string method, MouseButtons buttons, Point screen)
    { Point p = grip.PointToClient(screen); Event(grip, method, new MouseEventArgs(buttons, 1, p.X, p.Y, 0)); }
    static Point Start(Control grip) { return grip.PointToScreen(new Point(2, 2)); }
    static Point Delta(Point start, int x, int y) { return new Point(start.X + x, start.Y + y); }
    static bool Dragging(Control grip) { return (bool)grip.GetType().GetField("dragging", All).GetValue(grip); }
    static float Scale(Control control) { using (Graphics g = control.CreateGraphics()) return g.DpiX / 96f; }

    static void WidthTests(Assembly asm, string ns)
    {
        int saved = 0, savedWidth = 0;
        var ui = asm.GetType(ns + "Ui", true);
        var callback = ui.GetField("SaveControlWidth", All); object previous = callback.GetValue(null);
        callback.SetValue(null, new Action<Form, Control, int>(delegate(Form f, Control c, int width) { saved++; savedWidth = width; }));
        try
        {
            using (var form = new Form { ClientSize = new Size(500, 300), AutoScaleMode = AutoScaleMode.None })
            using (var input = new TextBox { Width = 153, Left = 20, Top = 20 })
            {
                form.Controls.Add(input); var grip = (Control)Make(asm.GetType(ns + "WidthGrip", true), input);
                IntPtr formHandle = form.Handle, handle = grip.Handle; Point start = Start(grip);
                Mouse(grip, "OnMouseDown", MouseButtons.Left, start);
                Check(grip.Capture && Dragging(grip), "width starts with real capture");
                Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, 37, 0));
                Check(input.Width == 190, "width screen delta is independent of a moving floating grip");
                grip.Capture = false;
                Check(!Dragging(grip) && input.Width == 153 && saved == 0, "lost width capture restores the user's starting width without saving");
                Mouse(grip, "OnMouseMove", MouseButtons.None, Delta(start, 90, 0)); Mouse(grip, "OnMouseUp", MouseButtons.Left, Delta(start, 90, 0));
                Check(input.Width == 153 && saved == 0, "width hover and late release after interruption do nothing");
                start = Start(grip); Mouse(grip, "OnMouseDown", MouseButtons.Left, start);
                Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, 27, 0)); Mouse(grip, "OnMouseUp", MouseButtons.Left, Delta(start, 27, 0));
                Check(input.Width == 180 && saved == 1 && !grip.Capture && !Dragging(grip), "normal width release commits once and does not roll back");
                Check(savedWidth == (int)Math.Round(180 / Scale(form)), "width persistence uses current 96 DPI logical pixels");
                Mouse(grip, "OnMouseUp", MouseButtons.Left, Delta(start, 27, 0));
                Check(saved == 1, "duplicate width release never commits twice");
                foreach (string signal in new string[] { "OnVisibleChanged", "RecreateHandle" })
                {
                    start = Start(grip); Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, 20, 0));
                    if (signal == "RecreateHandle") Event(grip, signal); else Event(grip, signal, EventArgs.Empty);
                    Check(input.Width == 180 && saved == 1 && !Dragging(grip), "width interruption restores previous user width: " + signal);
                }
                start = Start(grip); Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.None, Delta(start, 50, 0));
                Check(input.Width == 180 && saved == 1 && !grip.Capture, "width move without a left button cancels rather than resizing");
                start = Start(grip); Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, -500, 0));
                Check(input.Width == 40, "width retains its minimum"); Mouse(grip, "OnMouseUp", MouseButtons.Right, start);
                Check(input.Width == 180 && saved == 1, "non-left width release restores without saving");
                start = Start(grip); Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, 20, 0)); grip.Dispose();
                Check(input.Width == 180 && saved == 1, "disposing a width grip restores without saving");
                Check(!form.Visible, "width fixture never shows a form");
            }
            // Flow layout mode has the same transaction while the parent repositions children.
            using (var form = new Form { ClientSize = new Size(500, 300), AutoScaleMode = AutoScaleMode.None })
            using (var flow = new FlowLayoutPanel { Width = 450, Height = 100, WrapContents = false })
            using (var input = new TextBox { Width = 117 })
            {
                form.Controls.Add(flow); flow.Controls.Add(input); var grip = (Control)Make(asm.GetType(ns + "WidthGrip", true), input); flow.Controls.Add(grip);
                IntPtr handle = form.Handle; handle = grip.Handle; flow.PerformLayout(); Point start = Start(grip);
                Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, 19, 0));
                Check(input.Width == 136, "embedded width grip accepts the same screen delta"); grip.Capture = false;
                Check(input.Width == 117 && saved == 1, "embedded width cancellation preserves the original width");
            }
        }
        finally { callback.SetValue(null, previous); }
    }

    static void RowTests(Assembly asm, string ns, bool fixedBelow)
    {
        int saved = 0, savedHeight = 0;
        var callback = asm.GetType(ns + "Ui", true).GetField("SaveRowHeight", All); object previous = callback.GetValue(null);
        callback.SetValue(null, new Action<Form, string, int>(delegate(Form f, string key, int height) { saved++; savedHeight = height; }));
        try
        {
            using (var form = new Form { ClientSize = new Size(500, 500), AutoScaleMode = AutoScaleMode.None })
            using (var host = new TableLayoutPanel { Width = 350, Height = 350, RowCount = 3, ColumnCount = 1 })
            {
                form.Controls.Add(host); host.RowStyles.Add(new RowStyle(SizeType.Percent, 37)); host.RowStyles.Add(new RowStyle(SizeType.Absolute, 8)); host.RowStyles.Add(new RowStyle(SizeType.AutoSize, 23));
                var grip = (Control)Make(asm.GetType(ns + "RowGrip", true), host, 0, 2, "row.fixture", 40, 35, fixedBelow); host.Controls.Add(grip, 0, 1);
                IntPtr handle = form.Handle; handle = host.Handle; handle = grip.Handle; host.PerformLayout();
                int minimumAbove = (int)Math.Round(40 * Scale(grip)), minimumBelow = (int)Math.Round(35 * Scale(grip));
                int[] original = host.GetRowHeights(); Point start = Start(grip);
                Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, 0, 29));
                Check(host.RowStyles[0].SizeType == SizeType.Absolute && host.RowStyles[0].Height >= minimumAbove, "row drag converts measured height into a DPI-bounded preview");
                Check(!fixedBelow || (host.RowStyles[2].SizeType == SizeType.AutoSize && host.RowStyles[2].Height == 23), "fixed lower row retains its AutoSize style during dragging");
                grip.Capture = false;
                Check(host.RowStyles[0].SizeType == SizeType.Percent && host.RowStyles[0].Height == 37 && host.RowStyles[2].SizeType == SizeType.AutoSize && host.RowStyles[2].Height == 23 && saved == 0, "row capture loss restores both original styles without saving");
                Check(Equal(original, host.GetRowHeights()), "row cancellation restores layout geometry");
                Mouse(grip, "OnMouseMove", MouseButtons.None, Delta(start, 0, 80)); Mouse(grip, "OnMouseUp", MouseButtons.Left, Delta(start, 0, 80));
                Check(saved == 0 && host.RowStyles[0].SizeType == SizeType.Percent && host.RowStyles[0].Height == 37, "row hover and late release after interruption are inert");
                foreach (string signal in new string[] { "OnVisibleChanged", "RecreateHandle" })
                {
                    start = Start(grip); Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, 0, 15));
                    if (signal == "RecreateHandle") Event(grip, signal); else Event(grip, signal, EventArgs.Empty);
                    Check(host.RowStyles[0].SizeType == SizeType.Percent && host.RowStyles[0].Height == 37 && host.RowStyles[2].SizeType == SizeType.AutoSize && saved == 0 && !Dragging(grip), "row interruption restores styles: " + signal);
                }
                start = Start(grip); Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.None, Delta(start, 0, 80));
                Check(host.RowStyles[0].SizeType == SizeType.Percent && saved == 0 && !grip.Capture, "row move without a left button cancels");
                foreach (int amount in new int[] { 100000, -100000 })
                {
                    int[] before = host.GetRowHeights(); int total = before[0] + before[2];
                    start = Start(grip); Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, 0, amount));
                    if (fixedBelow)
                    {
                        Check(host.RowStyles[0].Height == (amount > 0 ? before[0] + amount : minimumAbove), "fixed lower row allows above growth and enforces its DPI minimum");
                        Check(host.RowStyles[2].SizeType == SizeType.AutoSize && host.RowStyles[2].Height == 23, "large fixed-row drag leaves lower style unchanged");
                    }
                    else
                    {
                        Check(host.RowStyles[0].Height == (amount > 0 ? total - minimumBelow : minimumAbove) && host.RowStyles[2].Height == (amount > 0 ? minimumBelow : total - minimumAbove), "large paired-row drag clamps at both DPI minimum boundaries");
                        int[] after = host.GetRowHeights();
                        Check(host.RowStyles[0].Height + host.RowStyles[2].Height == total && after[0] + after[2] == total, "paired-row preview preserves original total styles and geometry");
                    }
                    grip.Capture = false;
                    Check(saved == 0 && !Dragging(grip) && Equal(before, host.GetRowHeights()), "large drag rollback restores original geometry without persistence");
                }
                start = Start(grip); Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, 0, -1000));
                Check(host.RowStyles[0].Height == minimumAbove && (fixedBelow || host.RowStyles[2].Height >= minimumBelow), "row drag keeps both DPI minimum limits");
                Mouse(grip, "OnMouseUp", MouseButtons.Left, start); float completed = host.RowStyles[0].Height;
                Check(saved == 1 && completed == minimumAbove && !grip.Capture && !Dragging(grip), "normal row release commits once without undoing preview");
                Check(savedHeight == (int)Math.Round(completed / Scale(form)), "row persistence converts physical pixels to current logical DPI");
                Check(savedHeight >= 40 && (int)Math.Round(savedHeight * Scale(form)) >= minimumAbove, "persisted row minimum reads back at or above its 96-DPI logical minimum");
                Mouse(grip, "OnMouseUp", MouseButtons.Left, start); Check(saved == 1, "duplicate row release never commits twice");
                start = Start(grip); Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, 0, 30)); grip.Dispose();
                Check(host.RowStyles[0].SizeType == SizeType.Absolute && host.RowStyles[0].Height == completed && saved == 1, "disposing row grip preserves previously committed user height");
                Check(!form.Visible, "row fixture never shows a form");
            }
            if (!fixedBelow)
            {
                using (var form = new Form { ClientSize = new Size(100, 100), AutoScaleMode = AutoScaleMode.None })
                using (var host = new TableLayoutPanel { Width = 90, Height = 48, RowCount = 3, ColumnCount = 1 })
                {
                    form.Controls.Add(host); host.RowStyles.Add(new RowStyle(SizeType.Absolute, 20)); host.RowStyles.Add(new RowStyle(SizeType.Absolute, 8)); host.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
                    var grip = (Control)Make(asm.GetType(ns + "RowGrip", true), host, 0, 2, "row.too-small", 40, 35, false); host.Controls.Add(grip, 0, 1);
                    IntPtr handle = form.Handle; handle = host.Handle; handle = grip.Handle; host.PerformLayout(); int[] before = host.GetRowHeights(); int persisted = saved;
                    Check(before[0] + before[2] < (int)Math.Round(40 * Scale(grip)) + (int)Math.Round(35 * Scale(grip)), "small row fixture cannot fit DPI-converted minimums");
                    Point start = Start(grip); Mouse(grip, "OnMouseDown", MouseButtons.Left, start); Mouse(grip, "OnMouseMove", MouseButtons.Left, Delta(start, 0, 100000)); Mouse(grip, "OnMouseUp", MouseButtons.Left, start);
                    Check(!Dragging(grip) && !grip.Capture && saved == persisted && Equal(before, host.GetRowHeights()) && host.RowStyles[0].Height == 20 && host.RowStyles[2].Height == 20, "impossible paired minimums fail closed without changing layout or saving");
                }
            }
        }
        finally { callback.SetValue(null, previous); }
    }
    static bool Equal(int[] a, int[] b) { if (a.Length != b.Length) return false; for (int i = 0; i < a.Length; i++) if (a[i] != b[i]) return false; return true; }

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] static extern IntPtr CreateDesktop(string name, IntPtr device, IntPtr mode, uint flags, uint access, IntPtr attributes);
    [DllImport("user32.dll", SetLastError = true)] static extern bool SetThreadDesktop(IntPtr desktop);
    [DllImport("user32.dll")] static extern bool CloseDesktop(IntPtr desktop);
    [DllImport("user32.dll")] static extern IntPtr GetThreadDesktop(uint thread);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [STAThread] static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8; Exception failure = null; IntPtr desktop = IntPtr.Zero;
        try
        {
            if (args.Length != 1) throw new ArgumentException("GripProbe <addin.dll>");
            desktop = CreateDesktop("grip-fixture-" + Guid.NewGuid().ToString("N"), IntPtr.Zero, IntPtr.Zero, 0, 0x000F01FF, IntPtr.Zero);
            if (desktop == IntPtr.Zero) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Private desktop unavailable");
            var thread = new Thread(delegate()
            {
                try
                {
                    if (!SetThreadDesktop(desktop)) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Private desktop attachment failed");
                    Check(GetThreadDesktop(GetCurrentThreadId()) == desktop, "fixture capture belongs to its isolated desktop");
                    var asm = Assembly.LoadFrom(System.IO.Path.GetFullPath(args[0])); string ns = asm.GetName().Name + ".";
                    WidthTests(asm, ns); RowTests(asm, ns, false); RowTests(asm, ns, true);
                    Check(Application.OpenForms.Count == 0, "all hidden fixtures are disposed");
                }
                catch (Exception e) { failure = e; }
            });
            thread.SetApartmentState(ApartmentState.STA); thread.Start(); thread.Join();
            if (failure != null) throw failure;
            Console.WriteLine("GripProbe: " + checks + " PASS, 0 FAIL (hidden private desktop; no mouse input)"); return 0;
        }
        catch (Exception e) { Console.WriteLine("FAIL: " + e); return 1; }
        finally { if (desktop != IntPtr.Zero) CloseDesktop(desktop); }
    }
}
