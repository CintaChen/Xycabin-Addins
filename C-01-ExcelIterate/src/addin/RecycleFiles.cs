using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    internal static class RecycleFiles
    {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(string path, IntPtr context, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out IShellItem item);
        internal static void Send(string path)
        {
            IFileOperation operation = null; IShellItem item = null;
            try
            {
                operation = (IFileOperation)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("3AD05575-8857-4850-9277-11B85BDB8E09")));
                // Silent, no error UI, only specified item, recycle, fail rather than elevate.
                operation.SetOperationFlags(0x0004 | 0x0400 | 0x2000 | 0x00080000 | 0x00100000 | 0x00040000);
                Guid iid = new Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE");
                SHCreateItemFromParsingName(path, IntPtr.Zero, ref iid, out item);
                var sink = new RecycleSink(); operation.DeleteItem(item, sink); operation.PerformOperations();
                bool aborted; operation.GetAnyOperationsAborted(out aborted);
                if (aborted || !sink.Completed || sink.Error < 0 || File.Exists(path) || Directory.Exists(path)) throw new IOException("系统无法回收，文件保留；错误 0x" + sink.Error.ToString("X8"));
            }
            finally { if (item != null) Marshal.ReleaseComObject(item); if (operation != null) Marshal.ReleaseComObject(operation); }
        }
    }
    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellItem { }
    [ComImport, Guid("947AAB5F-0A5C-4C13-B4D6-4BF7836FC9F8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IFileOperation
    {
        void Advise(IFileOperationProgressSink sink, out uint cookie);
        void Unadvise(uint cookie);
        void SetOperationFlags(uint flags);
        void SetProgressMessage([MarshalAs(UnmanagedType.LPWStr)] string message);
        void SetProgressDialog([MarshalAs(UnmanagedType.Interface)] object dialog);
        void SetProperties([MarshalAs(UnmanagedType.Interface)] object properties);
        void SetOwnerWindow(uint hwnd);
        void ApplyPropertiesToItem(IShellItem item);
        void ApplyPropertiesToItems([MarshalAs(UnmanagedType.Interface)] object items);
        void RenameItem(IShellItem item, [MarshalAs(UnmanagedType.LPWStr)] string name, IFileOperationProgressSink sink);
        void RenameItems([MarshalAs(UnmanagedType.Interface)] object items, [MarshalAs(UnmanagedType.LPWStr)] string name);
        void MoveItem(IShellItem item, IShellItem destination, [MarshalAs(UnmanagedType.LPWStr)] string name, IFileOperationProgressSink sink);
        void MoveItems([MarshalAs(UnmanagedType.Interface)] object items, IShellItem destination);
        void CopyItem(IShellItem item, IShellItem destination, [MarshalAs(UnmanagedType.LPWStr)] string name, IFileOperationProgressSink sink);
        void CopyItems([MarshalAs(UnmanagedType.Interface)] object items, IShellItem destination);
        void DeleteItem(IShellItem item, IFileOperationProgressSink sink);
        void DeleteItems([MarshalAs(UnmanagedType.Interface)] object items);
        void NewItem(IShellItem destination, uint attributes, [MarshalAs(UnmanagedType.LPWStr)] string name, [MarshalAs(UnmanagedType.LPWStr)] string template, IFileOperationProgressSink sink);
        void PerformOperations();
        void GetAnyOperationsAborted([MarshalAs(UnmanagedType.Bool)] out bool aborted);
    }
    [ComVisible(true), Guid("04B0F1A7-9490-44BC-96E1-4296A31252E2"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IFileOperationProgressSink
    {
        [PreserveSig] int StartOperations(); [PreserveSig] int FinishOperations(int result);
        [PreserveSig] int PreRenameItem(uint flags, IShellItem item, [MarshalAs(UnmanagedType.LPWStr)] string name);
        [PreserveSig] int PostRenameItem(uint flags, IShellItem item, [MarshalAs(UnmanagedType.LPWStr)] string name, int result, IShellItem created);
        [PreserveSig] int PreMoveItem(uint flags, IShellItem item, IShellItem destination, [MarshalAs(UnmanagedType.LPWStr)] string name);
        [PreserveSig] int PostMoveItem(uint flags, IShellItem item, IShellItem destination, [MarshalAs(UnmanagedType.LPWStr)] string name, int result, IShellItem created);
        [PreserveSig] int PreCopyItem(uint flags, IShellItem item, IShellItem destination, [MarshalAs(UnmanagedType.LPWStr)] string name);
        [PreserveSig] int PostCopyItem(uint flags, IShellItem item, IShellItem destination, [MarshalAs(UnmanagedType.LPWStr)] string name, int result, IShellItem created);
        [PreserveSig] int PreDeleteItem(uint flags, IShellItem item);
        [PreserveSig] int PostDeleteItem(uint flags, IShellItem item, int result, IShellItem created);
        [PreserveSig] int PreNewItem(uint flags, IShellItem destination, [MarshalAs(UnmanagedType.LPWStr)] string name);
        [PreserveSig] int PostNewItem(uint flags, IShellItem destination, [MarshalAs(UnmanagedType.LPWStr)] string name, [MarshalAs(UnmanagedType.LPWStr)] string template, uint attributes, int result, IShellItem created);
        [PreserveSig] int UpdateProgress(uint total, uint done);
        [PreserveSig] int ResetTimer(); [PreserveSig] int PauseTimer(); [PreserveSig] int ResumeTimer();
    }
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    internal sealed class RecycleSink : IFileOperationProgressSink
    {
        internal bool Completed; internal int Error = unchecked((int)0x80004004);
        public int PreDeleteItem(uint flags, IShellItem item)
        {
            // The shell omits TSF_DELETE_RECYCLE_IF_POSSIBLE for non-recyclable items.
            return (flags & 0x80) != 0 ? 0 : unchecked((int)0x80004004);
        }
        public int PostDeleteItem(uint flags, IShellItem item, int result, IShellItem created) { Completed = true; Error = result; return 0; }
        public int StartOperations() { return 0; } public int FinishOperations(int result) { if (result < 0) Error = result; return 0; }
        public int PreRenameItem(uint f, IShellItem i, string n) { return 0; } public int PostRenameItem(uint f, IShellItem i, string n, int r, IShellItem c) { return 0; }
        public int PreMoveItem(uint f, IShellItem i, IShellItem d, string n) { return 0; } public int PostMoveItem(uint f, IShellItem i, IShellItem d, string n, int r, IShellItem c) { return 0; }
        public int PreCopyItem(uint f, IShellItem i, IShellItem d, string n) { return 0; } public int PostCopyItem(uint f, IShellItem i, IShellItem d, string n, int r, IShellItem c) { return 0; }
        public int PreNewItem(uint f, IShellItem d, string n) { return 0; } public int PostNewItem(uint f, IShellItem d, string n, string t, uint a, int r, IShellItem c) { return 0; }
        public int UpdateProgress(uint t, uint d) { return 0; } public int ResetTimer() { return 0; } public int PauseTimer() { return 0; } public int ResumeTimer() { return 0; }
    }
}
