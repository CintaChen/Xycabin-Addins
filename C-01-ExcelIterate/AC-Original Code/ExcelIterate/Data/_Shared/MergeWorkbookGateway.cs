using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Excel_Iterate
{
    /// <summary>一个已定位的工作簿 + 工作表，负责关闭自己打开的簿。</summary>
    internal sealed class MergeWorkbookHandle
    {
        internal object Book;
        internal object Sheet;
        internal bool OpenedByUs;
        internal string Path = "";
        internal string SheetName = "";
    }

    /// <summary>
    /// 工作簿定位与生命周期：Resolve 出来的 Book / Sheet 要跨「预检 → 执行 → 还原」
    /// 三次点击长期存活，因此从局部 ComBag 里 Detach 出来、登记到本对象的 held，
    /// 只在 Dispose 时逆序释放一次（Close 先于 Release，且绝不二次释放）。
    /// </summary>
    internal sealed class MergeWorkbookGateway : IDisposable
    {
        private readonly object app;
        private readonly List<object> opened = new List<object>();
        private readonly List<object> held = new List<object>();
        private bool disposed;

        internal MergeWorkbookGateway(object application) { app = application; }

        /// <summary>登记一个长期持有的引用（Resolve 里 Detach 出来的那些）。</summary>
        internal object Hold(object value)
        {
            if (value != null)
            {
                bool known = false;
                for (int i = 0; i < held.Count; i++) if (ReferenceEquals(held[i], value)) { known = true; break; }
                if (!known) held.Add(value);
            }
            return value;
        }

        /// <summary>已打开则复用，未打开则打开；源工作簿一律只读。</summary>
        internal MergeWorkbookHandle Resolve(string path, string sheetName, bool readOnly)
        {
            MergeWorkbookHandle handle = new MergeWorkbookHandle();
            handle.Path = path;
            string full;
            try { full = System.IO.Path.GetFullPath(path); }
            catch { throw new InvalidOperationException("工作簿路径无效：" + path); }

            object existing = FindOpen(full);
            if (existing != null)
            {
                handle.Book = Hold(existing);
                handle.OpenedByUs = false;
            }
            else
            {
                using (ComBag bag = new ComBag())
                {
                    dynamic books = bag.Own(((dynamic)app).Workbooks);
                    // 关闭「更新链接」等打开期提示：UpdateLinks=0（从不更新）、
                    // AskToUpdateLinks=false（不询问是否更新）、DisplayAlerts=false 兜底。
                    dynamic a = (dynamic)app;
                    bool alerts = ComFormulaHost.Boolean(a.DisplayAlerts, "DisplayAlerts");
                    bool ask = ComFormulaHost.Boolean(a.AskToUpdateLinks, "AskToUpdateLinks");
                    object book;
                    try
                    {
                        a.DisplayAlerts = false;
                        a.AskToUpdateLinks = false;
                        book = bag.Own(books.Open(path, 0, readOnly));   // 0 = xlUpdateLinksNever
                    }
                    finally { a.AskToUpdateLinks = ask; a.DisplayAlerts = alerts; }
                    if (book == null) throw new InvalidOperationException("无法打开工作簿：" + path);
                    handle.Book = Hold(bag.Detach(book));
                    handle.OpenedByUs = true;
                    opened.Add(handle.Book);
                }
            }
            handle.Sheet = Hold(ResolveSheet(handle.Book, sheetName));
            handle.SheetName = Convert.ToString(((dynamic)handle.Sheet).Name);
            return handle;
        }

        private object FindOpen(string fullPath)
        {
            using (ComBag bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)app).Workbooks);
                int count = Convert.ToInt32(books.Count);
                for (int i = 1; i <= count; i++)
                {
                    object book = null;
                    try
                    {
                        book = bag.Own(GetItem(books, i));
                        string name = Convert.ToString(((dynamic)book).FullName);
                        if (string.IsNullOrEmpty(name)) continue;
                        if (string.Equals(System.IO.Path.GetFullPath(name), fullPath, StringComparison.OrdinalIgnoreCase))
                            return bag.Detach(book);
                    }
                    catch { }
                }
            }
            return null;
        }

        /// <summary>列出工作簿的工作表名（供界面下拉；排除隐藏表）。</summary>
        internal static List<string> SheetNames(object book)
        {
            List<string> names = new List<string>();
            using (ComBag bag = new ComBag())
            {
                dynamic sheets = bag.Own(((dynamic)book).Worksheets);
                int count = Convert.ToInt32(sheets.Count);
                for (int i = 1; i <= count; i++)
                    using (ComBag one = new ComBag())
                    {
                        dynamic sheet = one.Own(GetItem(sheets, i));
                        string name = Convert.ToString(sheet.Name);
                        int visible = -1;
                        try { visible = Convert.ToInt32(sheet.Visible); }
                        catch { }
                        if (visible == -1) names.Add(name);
                    }
            }
            return names;
        }

        private object ResolveSheet(object book, string sheetName)
        {
            using (ComBag bag = new ComBag())
            {
                dynamic sheets = bag.Own(((dynamic)book).Worksheets);
                int count = Convert.ToInt32(sheets.Count);
                if (count == 0) throw new InvalidOperationException("工作簿没有工作表。");
                if (!string.IsNullOrEmpty(sheetName))
                {
                    for (int i = 1; i <= count; i++)
                    {
                        object sheet = bag.Own(GetItem(sheets, i));
                        string name = Convert.ToString(((dynamic)sheet).Name);
                        if (string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                            return Detach(bag, sheet);
                    }
                    throw new InvalidOperationException("找不到工作表「" + sheetName + "」。");
                }
                for (int i = 1; i <= count; i++)
                {
                    object sheet = bag.Own(GetItem(sheets, i));
                    int visible = -1;
                    try { visible = Convert.ToInt32(((dynamic)sheet).Visible); }
                    catch { }
                    if (visible == -1) return Detach(bag, sheet);
                }
                return Detach(bag, bag.Own(GetItem(sheets, 1)));
            }
        }

        // ComBag 释放时会 Release 它 Own 过的一切；交给调用方长期持有的引用必须先摘出来。
        private static object Detach(ComBag bag, object value) { return bag.Detach(value); }

        internal static dynamic GetItem(object collection, int index)
        {
            dynamic list = collection;
            try { return list.Item(index); }
            catch { return list[index]; }
        }

        /// <summary>关闭本功能自己打开的工作簿；用户已打开的一律不动。释放交给 Dispose，避免二次 Release。</summary>
        internal void CloseOpened()
        {
            List<string> failures = new List<string>();
            for (int i = opened.Count - 1; i >= 0; i--)
            {
                try { ((dynamic)opened[i]).Close(false); }
                catch (Exception ex) { failures.Add(ex.Message); }
            }
            opened.Clear();
            if (failures.Count > 0) throw new InvalidOperationException("关闭临时打开的工作簿失败：" + string.Join("；", failures.ToArray()));
        }

        /// <summary>先关自己打开的簿，再逆序释放全部长期持有的引用（每个只释放一次）。</summary>
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            List<string> failures = new List<string>();
            try { CloseOpened(); }
            catch (Exception ex) { failures.Add(ex.Message); }
            for (int i = held.Count - 1; i >= 0; i--)
            {
                try { if (Marshal2.IsCom(held[i])) Marshal2.Release(held[i]); }
                catch (Exception ex) { failures.Add(ex.Message); }
            }
            held.Clear();
            if (failures.Count > 0) throw new InvalidOperationException("释放工作簿引用失败：" + string.Join("；", failures.ToArray()));
        }
    }
}
