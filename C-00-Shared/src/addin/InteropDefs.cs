using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    // =====================================================================
    // COM 接口定义（不引用任何外部程序集，全部按 IID 手工声明）
    //
    // 所有 IID 均已在本机 WPS (12.1.0.28488) 二进制中取证验证：
    //   - IDTExtensibility2 : ksaddndr.dll (AddInDesignerObjects 类型库)
    //   - IRibbonExtensibility : ksoapi.dll (Office 类型库)
    // WPS 与 MS Office 使用完全相同的 IID，两套宿主均可加载。
    // =====================================================================

    // ext_ConnectMode: 0=AfterStartup 1=Startup 2=External 3=CommandLine 4=Solution 5=UISetup
    // ext_DisconnectMode: 0=HostShutdown 1=UserClosed 2=UISetupComplete 3=SolutionClosed
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [In][MarshalAs(UnmanagedType.IDispatch)] object Application,
            [In] int ConnectMode,
            [In][MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            [In] int RemoveMode,
            [In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [In][MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    // WPS 的 IRibbonExtensibility 为纯 dispinterface（GetCustomUI 的 dispid = 1）
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([In][MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    // =====================================================================
    // IPictureDisp —— getImage 回调的返回值类型（WPS 下透明兼容性最佳）
    // =====================================================================
    [ComImport]
    [Guid("7BF80980-BF32-101A-8BBB-00AA00300CAB")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPictureDisp
    {
    }

    // The public callback marker keeps its original IID/signature. Read IPicture's
    // IUnknown vtable internally: width/height are HIMETRIC, not bitmap pixels.
    [ComImport]
    [Guid("7BF80980-BF32-101A-8BBB-00AA00300CAB")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPictureMetrics
    {
        [PreserveSig] int GetHandle(out int value);
        [PreserveSig] int GetPalette(out int value);
        [PreserveSig] int GetPictureType(out short value);
        [PreserveSig] int GetWidth(out int value);
        [PreserveSig] int GetHeight(out int value);
    }

    internal static class RibbonImages
    {
        // OLE can cache the first DPI of the process. Measure its actual conversion;
        // neither current screen DPI nor changing the host awareness is reliable.
        private static readonly double[] Calibration = new double[2];

        internal static IPictureDisp Create(int logicalSize, Func<int, int, Bitmap> render, Action<string> log)
        {
            return (IPictureDisp)ConvertCore(logicalSize, render,
                delegate(Bitmap b) { return NativeOle.FromBitmap(b); }, ReadLogicalSize, Release, log, Calibration);
        }

        internal static SizeF ReadLogicalSize(object picture)
        {
            IPictureMetrics metrics = (IPictureMetrics)picture;
            int width, height;
            Marshal.ThrowExceptionForHR(metrics.GetWidth(out width));
            Marshal.ThrowExceptionForHR(metrics.GetHeight(out height));
            return new SizeF((float)(width * 96.0 / 2540.0), (float)(height * 96.0 / 2540.0));
        }

        internal static void Release(object picture)
        {
            if (picture != null && Marshal.IsComObject(picture)) Marshal.FinalReleaseComObject(picture);
        }

        // A narrow injectable seam allows offline tests to simulate OLE mappings and
        // failures without modifying process DPI or adding a production dependency.
        internal static object ConvertCore(int logicalSize, Func<int, int, Bitmap> render,
            Func<Bitmap, object> wrap, Func<object, SizeF> metrics, Action<object> release,
            Action<string> log, double[] calibration)
        {
            if (logicalSize != 16 && logicalSize != 32) throw new ArgumentOutOfRangeException("logicalSize");
            lock (calibration)
            {
                for (int attempt = 0; attempt < 2; attempt++)
                {
                    object picture = null;
                    try
                    {
                        if (calibration[0] <= 0 || calibration[1] <= 0)
                        {
                            object sample = null;
                            try
                            {
                                using (Bitmap bitmap = new Bitmap(32, 32, PixelFormat.Format24bppRgb))
                                    sample = wrap(bitmap);
                                SizeF measured = metrics(sample);
                                if (!Valid(measured.Width) || !Valid(measured.Height))
                                    throw new InvalidOperationException("Invalid OLE calibration dimensions");
                                calibration[0] = measured.Width / 32.0;
                                calibration[1] = measured.Height / 32.0;
                            }
                            finally { if (sample != null) release(sample); }
                        }
                        int width = PhysicalSize(logicalSize, calibration[0]);
                        int height = PhysicalSize(logicalSize, calibration[1]);
                        using (Bitmap bitmap = render(width, height)) picture = wrap(bitmap);
                        SizeF actual = metrics(picture);
                        if (!Valid(actual.Width) || !Valid(actual.Height)
                            || Math.Abs(actual.Width - logicalSize) > 1 || Math.Abs(actual.Height - logicalSize) > 1)
                            throw new InvalidOperationException("OLE image logical size " + actual + "; expected " + logicalSize);
                        object result = picture; picture = null;
                        return result; // Only the returned picture retains its OLE-owned HBITMAP.
                    }
                    catch (Exception ex)
                    {
                        SafeLog(log, "Ribbon image calibration attempt " + (attempt + 1) + ": " + ex.Message);
                        calibration[0] = calibration[1] = 0;
                    }
                    finally { if (picture != null) release(picture); }
                }
                SafeLog(log, "Ribbon image calibration failed; falling back to original " + logicalSize + "px conversion.");
                using (Bitmap bitmap = render(logicalSize, logicalSize)) return wrap(bitmap);
            }
        }

        private static int PhysicalSize(int logicalSize, double logicalPerPixel)
        {
            double pixels = logicalSize / logicalPerPixel;
            if (double.IsNaN(pixels) || double.IsInfinity(pixels) || pixels < 1 || pixels > 512)
                throw new InvalidOperationException("OLE image physical size outside supported range");
            return (int)Math.Round(pixels, MidpointRounding.AwayFromZero);
        }
        private static bool Valid(float value) { return value > 0 && !float.IsInfinity(value) && !float.IsNaN(value); }
        private static void SafeLog(Action<string> log, string text) { try { if (log != null) log(text); } catch { } }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PictDesc
    {
        public int cbSizeOfStruct;
        public int picType;
        public IntPtr hbitmap;
        public IntPtr hpal;
        public int unused;
    }

    internal static class NativeOle
    {
        [DllImport("oleaut32.dll", ExactSpelling = true, PreserveSig = false)]
        public static extern IPictureDisp OleCreatePictureIndirect(
            ref PictDesc pictdesc,
            [In] ref Guid iid,
            [MarshalAs(UnmanagedType.Bool)] bool fOwn);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr handle);

        internal static IPictureDisp FromBitmap(Bitmap bitmap)
        {
            IntPtr handle = bitmap.GetHbitmap();
            try
            {
                PictDesc pd = new PictDesc();
                pd.cbSizeOfStruct = Marshal.SizeOf(typeof(PictDesc)); pd.picType = 1;
                pd.hbitmap = handle;
                Guid iid = typeof(IPictureDisp).GUID;
                IPictureDisp picture = OleCreatePictureIndirect(ref pd, ref iid, true);
                if (picture == null) throw new InvalidOperationException("OLE returned no picture");
                handle = IntPtr.Zero; // fOwn=true transfers ownership only after success.
                return picture;
            }
            finally { if (handle != IntPtr.Zero) DeleteObject(handle); }
        }

        // 「查看链接」非模态窗找宿主窗口兜底用
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
    }
}
