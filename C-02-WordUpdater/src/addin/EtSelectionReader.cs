using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace WordUpdater
{
    internal sealed class EtSelectionInfo
    {
        public List<LinkPasteEngine.SourceArea> Areas = new List<LinkPasteEngine.SourceArea>();
        public string WorkbookPath;
        public string SheetName;
    }

    // One entry for each acquired COM reference, even if the RCW is reused.
    // Selection reading borrows its Application; it never owns or closes WPS.
    internal sealed class EtReferenceScope : IDisposable
    {
        private readonly List<object> acquired = new List<object>();
        private readonly Func<object, bool> isCom;
        private readonly Action<object> release;
        private readonly Action<string> report;
        private bool disposed;

        internal EtReferenceScope() : this(Marshal.IsComObject,
            value => { Marshal.ReleaseComObject(value); },
            message => Connect.Log("ET选区引用释放失败：" + message)) { }

        internal EtReferenceScope(Func<object, bool> checkCom, Action<object> releaseReference, Action<string> diagnostic)
        { isCom = checkCom; release = releaseReference; report = diagnostic; }

        internal object Own(object value)
        { if (value != null) acquired.Add(value); return value; }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            for (int i = acquired.Count - 1; i >= 0; i--)
                try { if (isCom(acquired[i])) release(acquired[i]); }
                catch (Exception error) { try { report(error.Message); } catch { } }
            acquired.Clear();
        }
    }

    // Clipboard LINK codes may retain only a bounding box. Read all running ET
    // candidates and let the caller match workbook, sheet and selected areas.
    // ROT/GetActiveObject can return an existing RCW: release only acquisitions,
    // never FinalRelease, CreateInstance, Quit or process termination.
    internal static class EtSelectionReader
    {
        private static readonly string[] ProgIds = { "KET.Application", "et.Application", "Excel.Application" };
        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx context);

        internal static List<EtSelectionInfo> TryReadAll(out string note)
        {
            note = null;
            List<EtSelectionInfo> result = new List<EtSelectionInfo>();
            string last = null;
            using (EtReferenceScope references = new EtReferenceScope())
            {
                List<object> apps = CollectEtApps(references);
                foreach (object app in apps)
                {
                    string why;
                    EtSelectionInfo info = ReadFrom(app, out why);
                    if (info != null) result.Add(info);
                    else if (why != null) last = why;
                }
            }
            note = result.Count > 0 ? null : (last ?? "未找到运行中的 WPS 表格。");
            return result;
        }

        private static List<object> CollectEtApps(EtReferenceScope references)
        {
            List<object> apps = new List<object>();
            List<string> guids = new List<string>();
            foreach (string id in ProgIds)
                try
                {
                    Type type = Type.GetTypeFromProgID(id);
                    if (type == null) continue;
                    string guid = type.GUID.ToString();
                    if (!guids.Contains(guid)) guids.Add(guid);
                }
                catch { }

            if (guids.Count > 0)
            {
                try { EnumerateRot(apps, guids, true, references); } catch { }
                if (apps.Count == 0)
                    try { EnumerateRot(apps, guids, false, references); } catch { }
            }
            foreach (string id in ProgIds)
                try { AddApp(apps, Marshal.GetActiveObject(id), references); }
                catch { }
            return apps;
        }

        private static void AddApp(List<object> apps, object value, EtReferenceScope references)
        {
            if (value == null) return;
            references.Own(value); // Duplicate discovery still acquired a reference.
            AddUnique(apps, value);
        }

        private static void AddUnique(List<object> apps, object value)
        {
            if (value == null) return;
            foreach (object existing in apps) if (ReferenceEquals(existing, value)) return;
            apps.Add(value);
        }

        private static void EnumerateRot(List<object> apps, List<string> guids, bool nameFilter, EtReferenceScope applications)
        {
            using (EtReferenceScope rotReferences = new EtReferenceScope())
            {
                IRunningObjectTable rot = null;
                int status;
                try { status = GetRunningObjectTable(0, out rot); }
                finally { rotReferences.Own(rot); }
                if (status != 0 || rot == null) return;
                IEnumMoniker enumerator = null;
                try { rot.EnumRunning(out enumerator); }
                finally { rotReferences.Own(enumerator); }
                if (enumerator == null) return;

                IMoniker[] item = new IMoniker[1];
                while (true)
                    using (EtReferenceScope itemReferences = new EtReferenceScope())
                    {
                        item[0] = null;
                        int next;
                        try { next = enumerator.Next(1, item, IntPtr.Zero); }
                        finally { itemReferences.Own(item[0]); }
                        if (next != 0 || item[0] == null) break;
                        IMoniker moniker = item[0];
                        string name = null;
                        try
                        {
                            IBindCtx context = null;
                            int created;
                            try { created = CreateBindCtx(0, out context); }
                            finally { itemReferences.Own(context); }
                            if (created == 0 && context != null) moniker.GetDisplayName(context, null, out name);
                        }
                        catch { }
                        if (nameFilter)
                        {
                            bool match = false;
                            if (name != null)
                                foreach (string guid in guids)
                                    if (name.IndexOf(guid, StringComparison.OrdinalIgnoreCase) >= 0) { match = true; break; }
                            if (!match) continue;
                        }
                        object application = null;
                        try
                        {
                            try { rot.GetObject(moniker, out application); }
                            finally { applications.Own(application); }
                            AddUnique(apps, application);
                        }
                        catch { }
                    }
            }
        }

        private static EtSelectionInfo ReadFrom(object application, out string note)
        { return ReadFrom(application, out note, new EtReferenceScope()); }

        // The injected scope exercises the actual read/early-return lifecycle in
        // offline tests. The passed Application is borrowed and is not registered.
        private static EtSelectionInfo ReadFrom(object application, out string note, EtReferenceScope references)
        {
            note = null;
            using (references)
            {
                if (application == null) { note = "表格实例为空。"; return null; }
                dynamic app = application;
                dynamic selection = null;
                try { selection = references.Own((object)app.Selection); } catch { }
                if (selection == null) { note = "WPS 表格无可用选区。"; return null; }
                string address = null;
                try { address = (string)selection.Address; } catch { }
                if (address == null) { note = "当前选区不是单元格区域。"; return null; }

                dynamic areas = null;
                int count = 1;
                try
                {
                    areas = references.Own((object)selection.Areas);
                    count = (int)areas.Count;
                }
                catch { count = 1; }
                if (count <= 1) { note = "选区只有一片，无需逐片明细。"; return null; }

                EtSelectionInfo result = new EtSelectionInfo();
                for (int i = 1; i <= count; i++)
                {
                    dynamic area = null;
                    try { area = references.Own((object)UpdateEngine.GetItem(areas, i)); } catch { }
                    if (area == null) { note = "读取子区域 " + i + " 失败。"; return null; }
                    string areaAddress = null, sheetName = null, workbookPath = null;
                    dynamic sheet = null, book = null;
                    try { areaAddress = (string)area.Address; } catch { }
                    try { sheet = references.Own((object)area.Worksheet); } catch { }
                    if (sheet != null)
                    {
                        try { sheetName = (string)sheet.Name; } catch { }
                        try { book = references.Own((object)sheet.Parent); } catch { }
                    }
                    if (book != null) try { workbookPath = (string)book.FullName; } catch { }
                    if (areaAddress == null) { note = "子区域 " + i + " 无地址。"; return null; }
                    result.WorkbookPath = workbookPath;
                    result.SheetName = sheetName;
                    LinkPasteEngine.SourceArea parsed = ParseAddress(areaAddress, sheetName);
                    if (parsed == null) { note = "无法解析地址 " + areaAddress + "。"; return null; }
                    result.Areas.Add(parsed);
                }
                if (result.Areas.Count < 2) { note = "未取到多个子区域。"; return null; }
                return result;
            }
        }

        private static LinkPasteEngine.SourceArea ParseAddress(string address, string sheetName)
        {
            List<LinkPasteEngine.SourceArea> areas = new List<LinkPasteEngine.SourceArea>();
            string error;
            if (!LinkPasteEngine.TrySplitItem(address, areas, out error) || areas.Count != 1) return null;
            LinkPasteEngine.SourceArea area = areas[0];
            area.SheetPart = string.IsNullOrEmpty(sheetName) ? "" : sheetName + "!";
            return area;
        }
    }
}
