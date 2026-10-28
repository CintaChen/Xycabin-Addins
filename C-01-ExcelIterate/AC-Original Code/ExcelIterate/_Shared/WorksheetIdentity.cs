using System;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    // The caller owns the worksheet reference; this snapshot owns only Parent.
    // CodeName belongs to the VBA project and can legitimately be empty in WPS.
    internal sealed class WorksheetIdentity : IDisposable
    {
        private readonly object worksheet;
        private object workbook;
        internal readonly string Path, Name;
        private bool disposed;
        private WorksheetIdentity(object sheet)
        {
            worksheet = sheet;
            try
            {
                if (sheet == null) throw new InvalidOperationException("工作表不可用。");
                workbook = ((dynamic)sheet).Parent;
                Path = workbook == null ? null : ((dynamic)workbook).FullName as string;
                Name = ((dynamic)sheet).Name as string;
                if (string.IsNullOrEmpty(Path) || string.IsNullOrEmpty(Name))
                    throw new InvalidOperationException("无法核对工作簿路径或工作表名称。");
            }
            catch { Dispose(); throw; }
        }
        internal static WorksheetIdentity Capture(object worksheet) { return new WorksheetIdentity(worksheet); }
        internal static bool SameObject(object first, object second)
        {
            if (first == null || second == null) return false;
            if (ReferenceEquals(first, second)) return true;
            if (!Marshal.IsComObject(first) || !Marshal.IsComObject(second)) return false;
            IntPtr a = IntPtr.Zero, b = IntPtr.Zero;
            try
            {
                a = Marshal.GetIUnknownForObject(first); b = Marshal.GetIUnknownForObject(second);
                return a == b;
            }
            finally
            {
                if (b != IntPtr.Zero) Marshal.Release(b);
                if (a != IntPtr.Zero) Marshal.Release(a);
            }
        }
        internal bool Matches(object candidate)
        {
            if (disposed) throw new ObjectDisposedException("WorksheetIdentity");
            if (!SameObject(worksheet, candidate)) return false;
            object parent = ((dynamic)candidate).Parent;
            try
            {
                return SameObject(workbook, parent)
                    && string.Equals(Path, ((dynamic)parent).FullName as string, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(Name, ((dynamic)candidate).Name as string, StringComparison.Ordinal);
            }
            finally { Release(parent); }
        }
        private static void Release(object value)
        { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        public void Dispose()
        {
            if (disposed) return;
            disposed = true; object parent = workbook; workbook = null; Release(parent);
        }
    }
}
