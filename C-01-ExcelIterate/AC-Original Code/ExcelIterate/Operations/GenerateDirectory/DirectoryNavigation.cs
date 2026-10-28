using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;

namespace Excel_Iterate
{
    [ComVisible(true), Guid("00024413-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDirectoryAppEvents
    {
        [DispId(1854)] void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sheet, [In, MarshalAs(UnmanagedType.IDispatch)] object hyperlink);
        [DispId(1564)] void SheetChange([In,MarshalAs(UnmanagedType.IDispatch)] object sheet,[In,MarshalAs(UnmanagedType.IDispatch)] object target);
        [DispId(1563)] void SheetCalculate([In,MarshalAs(UnmanagedType.IDispatch)] object sheet);
        [DispId(1570)] void WorkbookBeforeClose([In,MarshalAs(UnmanagedType.IDispatch)] object book,[In,Out] ref bool cancel);
        [DispId(1573)] void WorkbookNewSheet([In,MarshalAs(UnmanagedType.IDispatch)] object book,[In,MarshalAs(UnmanagedType.IDispatch)] object sheet);
        // Verified against the installed Microsoft Excel 15.0 PIA metadata.
        [DispId(3079)] void SheetBeforeDelete([In,MarshalAs(UnmanagedType.IDispatch)] object sheet);
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public sealed class DirectoryEventSink : IDirectoryAppEvents
    {
        internal object App;
        // Reuse the existing public COM event contract; scanning adds no public types.
        internal Action<object,bool> SourceChanged;
        public void SheetChange(object sheet,object target){if(SourceChanged!=null)SourceChanged(sheet,false);}
        public void SheetCalculate(object sheet){if(SourceChanged!=null)SourceChanged(sheet,false);}
        public void WorkbookBeforeClose(object book,ref bool cancel){if(SourceChanged!=null)SourceChanged(book,true);}
        public void WorkbookNewSheet(object book,object sheet){if(SourceChanged!=null)SourceChanged(book,true);}
        public void SheetBeforeDelete(object sheet){if(SourceChanged!=null)SourceChanged(sheet,false);}
        private bool handling;
        public void SheetFollowHyperlink(object sheet, object hyperlink)
        {
            if (handling || App == null) return;
            bool ours = false; handling = true;
            try
            {
                using (ComBag bag = new ComBag())
                {
                    dynamic book = bag.Own(((dynamic)sheet).Parent), names = bag.Own(book.Names);
                    dynamic clicked = bag.Own(((dynamic)hyperlink).Range);
                    for (int i = 1; i <= Convert.ToInt32(names.Count); i++)
                    {
                        using (ComBag one = new ComBag())
                        {
                            dynamic name = one.Own(names.Item(i)); string key = (string)name.Name;
                            int prefix = key.IndexOf("_EI_Dir_A_", StringComparison.Ordinal);
                            if (prefix < 0) continue;
                            dynamic anchor, anchorSheet;
                            // Deleted older directory sheets leave orphan names;
                            // these must not block a newer directory's hyperlinks.
                            try { anchor = one.Own(name.RefersToRange); anchorSheet = one.Own(anchor.Worksheet); } catch { continue; }
                            if (!WorksheetIdentity.SameObject((object)anchorSheet, sheet) || (int)anchor.Row != (int)clicked.Row || (int)anchor.Column != (int)clicked.Column) continue;
                            ours = true;
                            dynamic targetName = one.Own(names.Item(key.Substring(prefix).Replace("_EI_Dir_A_", "_EI_Dir_T_")));
                            dynamic target = one.Own(targetName.RefersToRange), targetSheet = one.Own(target.Worksheet);
                            if (ComFormulaHost.Boolean((object)book.ProtectStructure, "ProtectStructure") && Convert.ToInt32(targetSheet.Visible) != -1)
                                throw new InvalidOperationException("工作簿结构受保护，不能显示隐藏工作表；请先取消结构保护。");
                            int originalVisibility = Convert.ToInt32(targetSheet.Visible);
                            try { targetSheet.Visible = -1; targetSheet.Activate(); ((dynamic)App).Goto(target, false); }
                            catch { if (originalVisibility != -1) { ((dynamic)sheet).Activate(); targetSheet.Visible = originalVisibility; } throw; }
                            return;
                        }
                    }
                }
            }
            catch (Exception ex) { Connect.Log("目录跳转：" + ex); if (ours) UiDialog.Show(null, "无法跳转：" + ex.Message + "\r\n目标已删除时请重新生成目录。", "生成目录", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { handling = false; }
        }
    }

    internal sealed class DirectoryNavigation : IDisposable
    {
        private IConnectionPoint point;
        private DirectoryEventSink sink;
        private int cookie;
        internal DirectoryNavigation(object app)
        {
            IConnectionPointContainer container = app as IConnectionPointContainer;
            if (container == null) throw new InvalidOperationException("当前宿主不支持目录超链接事件。");
            Guid id = new Guid("00024413-0000-0000-C000-000000000046");
            try { container.FindConnectionPoint(ref id, out point); sink = new DirectoryEventSink { App = app }; point.Advise(sink, out cookie); }
            catch { Dispose(); throw; }
        }
        public void Dispose()
        {
            if (sink != null) sink.App = null;
            try { if (point != null && cookie != 0) point.Unadvise(cookie); }
            finally { cookie = 0; if (point != null && Marshal.IsComObject(point)) Marshal.ReleaseComObject(point); point = null; sink = null; }
        }
        internal static int Generate(object app)
        {
            using (ComBag bag = new ComBag())
            {
                dynamic book = bag.Own(((dynamic)app).ActiveWorkbook);
                if (book == null) throw new InvalidOperationException("请先打开工作簿。");
                dynamic sheets = bag.Own(book.Worksheets), names = bag.Own(book.Names);
                List<object> targets = new List<object>();
                for (int i = 1; i <= Convert.ToInt32(sheets.Count); i++) targets.Add(bag.Own(sheets.Item(i)));
                object output = null; List<object> createdNames = new List<object>(); string token = Guid.NewGuid().ToString("N"); bool success = false;
                using (HostQuiet quiet = new HostQuiet(app)) try
                {
                    output = WorkbookUi.NewSheet((object)book, "目录", bag, true);
                    object[,] values = new object[targets.Count + 1, 3]; values[0, 0] = "工作表名称"; values[0, 1] = "跳转"; values[0, 2] = "生成时状态";
                    for (int i = 0; i < targets.Count; i++) { dynamic s = targets[i]; values[i + 1, 0] = s.Name; values[i + 1, 1] = "跳转至本表"; int v = Convert.ToInt32(s.Visible); values[i + 1, 2] = v == -1 ? "可见" : v == 2 ? "非常隐藏" : "隐藏"; }
                    dynamic range = bag.Own(((dynamic)output).Range["A1:C" + (targets.Count + 1)]); range.NumberFormat = "@"; range.Value2 = values;
                    dynamic links = bag.Own(((dynamic)output).Hyperlinks);
                    for (int i = 0; i < targets.Count; i++)
                    {
                        dynamic s = targets[i]; string index = token + "_" + i, anchorAddress = "$B$" + (i + 2);
                        string destination = WorkbookUi.Quote((string)s.Name) + "!$A$1";
                        object targetName = bag.Own(names.Add("_EI_Dir_T_" + index, "=" + destination, false)); createdNames.Add(targetName);
                        object anchorName = bag.Own(names.Add("_EI_Dir_A_" + index, "=" + WorkbookUi.Quote((string)((dynamic)output).Name) + "!" + anchorAddress, false)); createdNames.Add(anchorName);
                        dynamic cell = bag.Own(((dynamic)output).Range[anchorAddress]);
                        // A valid self target triggers the event without the host's
                        // invalid-hidden-target warning; the sink navigates afterward.
                        string address = Convert.ToInt32(s.Visible) == -1 ? destination : WorkbookUi.Quote((string)((dynamic)output).Name) + "!" + anchorAddress;
                        bag.Own(links.Add(cell, "", address, "跳转至 " + (string)s.Name, "跳转至本表"));
                    }
                    WorkbookUi.StyleTable(output, 3, targets.Count + 1);
                    // Apply after Hyperlinks.Add and the shared table style so every
                    // directory cell, including clickable links, retains this format.
                    dynamic directoryFont = bag.Own(range.Font);
                    directoryFont.Bold = true;
                    range.HorizontalAlignment = -4108; // xlCenter
                    range.VerticalAlignment = -4108;
                    WorkbookUi.FitOutput(output, 1, 3, targets.Count + 1);
                    ((dynamic)output).Activate(); success = true; return targets.Count;
                }
                finally
                {
                    if (!success)
                    {
                        List<string> failures = new List<string>();
                        foreach (object name in createdNames) try { ((dynamic)name).Delete(); } catch (Exception ex) { failures.Add(ex.Message); }
                        if (output != null) try { WorkbookUi.DeleteSheet(app, output); } catch (Exception ex) { failures.Add(ex.Message); }
                        if (failures.Count > 0) throw new InvalidOperationException("目录清理不完整：" + string.Join("；", failures));
                    }
                }
            }
        }
    }
}
