using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Collections;
using System.Reflection;
using WordUpdater;

// 实机测试必须显式运行。Application 可能是用户已有单实例；只关闭自建样本，
// 不退出宿主、不扫描或结束进程。0=PASS，1=FAIL，2=SKIP。
public sealed class NoSelectionEtHost
{
    private readonly object application;
    public NoSelectionEtHost(object value) { application = value; }
    public object Selection { get { return null; } }
    public object Workbooks { get { return ((dynamic)application).Workbooks; } }
}

internal static class LinkPasteTest
{
    private const int Rows = 4, Cols = 4;
    private static readonly string[] ExcelIds = { "KET.Application", "et.Application", "Excel.Application" };
    private static readonly string[] WordIds = { "KWPS.Application", "wps.Application", "Word.Application" };

    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        string export = null;
        try
        {
            if (args.Length > 1) throw new ArgumentException("Usage: LinkPasteTest [new-source-export.xlsx]");
            if (args.Length == 1)
            {
                export = Path.GetFullPath(args[0]);
                ProbeSession.RejectExistingOutput(export);
            }
            using (ProbeSession session = new ProbeSession("link-paste-test"))
            {
                int result = 1;
                object document = null;
                try
                {
                    string sourcePath = session.Output("source.xlsx");
                    object et = session.Take(ProbeSession.Connect(ExcelIds, false));
                    object wps = session.Take(ProbeSession.Connect(WordIds, false));
                    if (et == null || wps == null)
                    {
                        Console.WriteLine("SKIP: WPS 表格或文字 COM 宿主不可用。");
                        result = 2;
                    }
                    else
                    {
                        dynamic excel = et, word = wps;
                        session.Override(() => (object)excel.Visible, value => excel.Visible = value, true, "Excel.Visible");
                        session.Override(() => (object)excel.DisplayAlerts, value => excel.DisplayAlerts = value, false, "Excel.DisplayAlerts");
                        session.Override(() => (object)word.Visible, value => word.Visible = value, true, "Word.Visible");
                        session.Override(() => (object)word.DisplayAlerts, value => word.DisplayAlerts = value, 0, "Word.DisplayAlerts");

                        dynamic books = session.Take((object)excel.Workbooks);
                        object workbook = books.Add();
                        session.Own(workbook, () => ((dynamic)workbook).Close(false));
                        dynamic book = workbook;
                        dynamic sheets = session.Take((object)book.Worksheets);
                        object sheet = session.Take((object)sheets[1]);
                        for (int row = 1; row <= Rows; row++)
                            for (int col = 1; col <= Cols; col++)
                                SetCell(session, sheet, row, col, "R" + row + "C" + col);
                        book.SaveAs(sourcePath);
                        CopyRange(session, sheet, "A1:D4");

                        dynamic documents = session.Take((object)word.Documents);
                        document = documents.Add();
                        session.Own(document, () => ((dynamic)document).Close(0));
                        dynamic doc = document;
                        dynamic start = session.Take((object)doc.Range(0, 0));
                        dynamic tables = session.Take((object)doc.Tables);
                        object table = session.Take((object)tables.Add(start, Rows, Cols));
                        SelectTable(session, table);
                        int before = Convert.ToInt32(documents.Count);
                        LinkPasteResult linked = LinkPasteEngine.Execute(wps);
                        PrintLink("整表", linked);
                        Check(linked.Status == LinkPasteStatus.Completed, linked.ErrorMessage);
                        Check(linked.CellsLinked == Rows * Cols, "整表链接格数应为16");
                        Check(linked.TargetRows == Rows && linked.TargetCols == Cols && linked.TargetCells == Rows * Cols, "目标矩形应为4×4");
                        Check(linked.SourceCells == Rows * Cols, "源序列应为16格");
                        Check(Convert.ToInt32(documents.Count) == before, "链接引擎应无残留临时文档");
                        CheckGrid(session, table, "R");
                        using (ProbeReferences refs = session.References())
                        {
                            dynamic tableRange = refs.Take((object)((dynamic)table).Range);
                            dynamic fields = refs.Take((object)tableRange.Fields);
                            dynamic content = refs.Take((object)doc.Content);
                            dynamic allFields = refs.Take((object)content.Fields);
                            Check(Convert.ToInt32(fields.Count) == Rows * Cols, "表格字段数应为16");
                            Check(Convert.ToInt32(allFields.Count) == Rows * Cols, "文档字段数应为16");
                            dynamic tail = refs.Take((object)doc.Range(Convert.ToInt32(tableRange.End) - 1, Convert.ToInt32(content.End)));
                            Check(Convert.ToString(tail.Text).Replace("\r", "").Replace("\a", "").Trim().Length == 0, "文末应无探测残留文字");
                        }

                        for (int row = 1; row <= Rows; row++)
                            for (int col = 1; col <= Cols; col++)
                                SetCell(session, sheet, row, col, "U" + row + "C" + col);
                        book.Save();
                        SelectTable(session, table);
                        RunResult updated = UpdateEngine.Execute(wps);
                        Console.WriteLine("UpdateEngine: " + updated.Status + " found=" + updated.Found + " ok=" + updated.Success + " fail=" + updated.Failed);
                        Check(updated.Status == RunStatus.Completed && updated.Success == Rows * Cols, "批量更新16格：" + updated.ErrorMessage);
                        CheckGrid(session, table, "U");

                        CopyRange(session, sheet, "A1:B2");
                        SelectRect(session, document, table, 1, 1, 2, 2);
                        LinkPasteResult rectangle = LinkPasteEngine.Execute(wps);
                        PrintLink("子矩形", rectangle);
                        Check(rectangle.Status == LinkPasteStatus.Completed, rectangle.ErrorMessage);
                        Check(rectangle.TargetRows == 2 && rectangle.TargetCols == 2 && rectangle.SourceCells == 4 && rectangle.CellsLinked == 4, "子矩形应为2×2/4格");
                        for (int row = 1; row <= 2; row++)
                            for (int col = 1; col <= 2; col++)
                                Check(CellText(session, table, row, col) == "U" + row + "C" + col, "子矩形按行优先映射");

                        SetCell(session, sheet, 1, 8, "XH1");
                        SetCell(session, sheet, 3, 8, "XH3");
                        book.Save();
                        CopyRange(session, sheet, "H1,H3"); // 先Select再Copy，供路径B回读实时选区。
                        using (ProbeReferences refs = session.References())
                        {
                            dynamic selection = refs.Take((object)excel.Selection);
                            dynamic areas = refs.Take((object)selection.Areas);
                            Check(Convert.ToInt32(areas.Count) == 2, "Ctrl多区域源应为2片");
                        }
                        SelectRect(session, document, table, 1, 1, 2, 1);
                        LinkPasteResult multiple = LinkPasteEngine.Execute(wps);
                        PrintLink("多区域", multiple);
                        Check(multiple.Status == LinkPasteStatus.Completed, multiple.ErrorMessage);
                        Check(multiple.TargetRows == 2 && multiple.TargetCols == 1 && multiple.TargetCells == 2 && multiple.SourceCells == 2 && multiple.CellsLinked == 2, "多区域应映射2格");
                        Check(CellText(session, table, 1, 1) == "XH1" && CellText(session, table, 2, 1) == "XH3", "路径B回读H1/H3");
                        CopyRange(session, sheet, "A2:D2,A4:D4");
                        CheckWindowMultiSelection(et, Convert.ToString(book.FullName));
                        SelectRect(session, document, table, 1, 1, 2, 4);
                        CheckSelection(session, wps, "1,1;1,2;1,3;1,4;2,1;2,2;2,3;2,4");
                        LinkPasteResult rows = LinkPasteEngine.Execute(wps); PrintLink("Ctrl第2/4行到2x4", rows);
                        Check(rows.Status == LinkPasteStatus.Completed && rows.SourceCells == 8 && rows.CellsLinked == 8
                            && rows.Refreshed == 8 && rows.TargetRows == 2 && rows.TargetCols == 4,
                            "第2/4行只建立8条链接，不计入第3行：" + rows.ErrorMessage);
                        for (int row = 1; row <= 2; row++)
                            for (int col = 1; col <= 4; col++)
                            {
                                int sourceRow = row == 1 ? 2 : 4;
                                Check(CellText(session, table, row, col) == "U" + sourceRow + "C" + col, "两组整行数据按顺序对应");
                                CheckReference(session, table, row, col, "R" + sourceRow + "C" + col);
                            }
                        CheckMergedHeader(session, wps, sheet, workbook);
                        if (export != null) session.Export(sourcePath, export);
                        result = 0;
                    }
                }
                catch (Exception error)
                {
                    session.Fail(error.ToString());
                    if (document != null) PreserveDocument(session, document);
                    result = 1;
                }
                int exit = session.Finish(result);
                PrintLogTail();
                return exit;
            }
        }
        catch (Exception error)
        {
            Console.WriteLine("FAIL: " + error);
            return 1;
        }
    }

    private static void Check(bool passed, string message)
    {
        if (!passed) throw new InvalidOperationException("校验失败：" + message);
        Console.WriteLine("PASS: " + message);
    }

    private static void CheckMergedHeader(ProbeSession session, object application, object sheet, object workbook)
    {
        object document = null;
        try
        {
            using (ProbeReferences refs = session.References())
            {
                dynamic docs = refs.Take(((dynamic)application).Documents);
                document = docs.Add(); session.Own(document, () => ((dynamic)document).Close(0));
                dynamic doc = document, tables = refs.Take(doc.Tables), start = refs.Take(doc.Range(0, 0));
                object table = session.Take((object)tables.Add(start, 3, 4));
                ((dynamic)table).AllowAutoFit = false;
                dynamic head = refs.Take(((dynamic)table).Cell(1, 1)), end = refs.Take(((dynamic)table).Cell(1, 4));
                head.Merge(end);
                SetWordCell(session, doc, table, 1, 1, "合并表头");
                for (int row = 2; row <= 3; row++)
                    for (int col = 1; col <= 4; col++) SetWordCell(session, doc, table, row, col, "保留" + row + col);
                dynamic headerRange = refs.Take(head.Range), font = refs.Take(headerRange.Font);
                font.Bold = -1; font.Size = 14; headerRange.HighlightColorIndex = 3;
                Check(!Convert.ToBoolean(((dynamic)table).Uniform), "样本整表确实存在合并表头");
                string kept = OutsideSignature(session, table), structure = Structure(session, table);
                for (int row = 1; row <= 2; row++)
                    for (int col = 1; col <= 2; col++) SetCell(session, sheet, row, col, "M" + row + "C" + col);
                ((dynamic)workbook).Save(); CopyRange(session, sheet, "A1:B2");
                SelectRect(session, document, table, 2, 1, 3, 2);
                CheckSelection(session, application, "2,1;2,2;3,1;3,2");
                int beforeDocuments = Convert.ToInt32(docs.Count);
                LinkPasteResult linked = LinkPasteEngine.Execute(application); PrintLink("合并表头下四格", linked);
                Check(linked.Status == LinkPasteStatus.Completed && linked.CellsLinked == 4 && linked.Refreshed == 4,
                    "合并表头下四格均已链接并刷新：" + linked.ErrorMessage);
                Check(linked.TargetRow1 == 2 && linked.TargetCol1 == 1 && linked.TargetRows == 2 && linked.TargetCols == 2,
                    "目标严格为第2~3行、第1~2列");
                for (int row = 1; row <= 2; row++)
                    for (int col = 1; col <= 2; col++)
                    {
                        Check(CellText(session, table, row + 1, col) == "M" + row + "C" + col, "合并表头下显示值按行优先对应");
                        CheckReference(session, table, row + 1, col, "R" + row + "C" + col);
                    }
                dynamic content = refs.Take(doc.Content), fields = refs.Take(content.Fields);
                Check(Convert.ToInt32(fields.Count) == 4 && Convert.ToInt32(docs.Count) == beforeDocuments,
                    "文档只有四条目标字段且无临时文档残留");
                Check(OutsideSignature(session, table) == kept && Structure(session, table) == structure,
                    "合并表头和其他单元格内容、格式、表格结构保持原样");
                for (int row = 1; row <= 2; row++)
                    for (int col = 1; col <= 2; col++) SetCell(session, sheet, row, col, "N" + row + "C" + col);
                ((dynamic)workbook).Save(); SelectRect(session, document, table, 2, 1, 3, 2);
                RunResult updated = UpdateEngine.Execute(application);
                Check(updated.Status == RunStatus.Completed && updated.Success == 4, "合并表头下原四条链接可批量更新");
                for (int row = 1; row <= 2; row++)
                    for (int col = 1; col <= 2; col++)
                        Check(CellText(session, table, row + 1, col) == "N" + row + "C" + col, "源变化后四格显示新值");
                Check(OutsideSignature(session, table) == kept, "更新链接也保留选区外内容和格式");

                head.Select();
                CheckRejectedUnchanged(session, application, doc, LinkPasteStatus.SelectionNotRect, "单独选中合并表头");
                CopyRange(session, sheet, "A1"); SelectRect(session, document, table, 2, 1, 3, 2);
                CheckRejectedUnchanged(session, application, doc, LinkPasteStatus.SizeMismatch, "源目标格数不匹配");

                // 左侧合并会改变 WPS 的物理 ColumnIndex；右侧格子仍应按网格对齐。
                object secondDocument = docs.Add(); session.Own(secondDocument, () => ((dynamic)secondDocument).Close(0));
                dynamic second = secondDocument, secondTables = refs.Take(second.Tables), origin = refs.Take(second.Range(0, 0));
                object shifted = session.Take((object)secondTables.Add(origin, 2, 4));
                dynamic left = refs.Take(((dynamic)shifted).Cell(1, 1)), right = refs.Take(((dynamic)shifted).Cell(1, 2));
                left.Merge(right); SetWordCell(session, second, shifted, 1, 1, "左侧合并保留");
                CopyRange(session, sheet, "A1:B2"); SelectRect(session, second, shifted, 1, 2, 2, 4);
                CheckSelection(session, application, "1,2;1,3;2,3;2,4");
                LinkPasteResult shiftedResult = LinkPasteEngine.Execute(application); PrintLink("左侧合并右侧四格", shiftedResult);
                Check(shiftedResult.Status == LinkPasteStatus.Completed && shiftedResult.CellsLinked == 4
                    && shiftedResult.TargetCol1 == 3 && shiftedResult.TargetCols == 2, "左侧合并不导致错列或误拒绝");
                Check(CellText(session, shifted, 1, 1) == "左侧合并保留", "左侧未选中合并格保持原样");
                Check(CellText(session, shifted, 1, 2) == "N1C1" && CellText(session, shifted, 1, 3) == "N1C2"
                    && CellText(session, shifted, 2, 3) == "N2C1" && CellText(session, shifted, 2, 4) == "N2C2",
                    "物理列号不一致的四格链接位置与内容均正确");
            }
        }
        catch { if (document != null) PreserveDocument(session, document); throw; }
    }

    private static void CheckWindowMultiSelection(object application, string sourcePath)
    {
        Type reader = typeof(LinkPasteEngine).Assembly.GetType("WordUpdater.EtSelectionReader", true);
        Type scope = reader.Assembly.GetType("WordUpdater.EtReferenceScope", true);
        MethodInfo read = reader.GetMethod("ReadCandidatesFrom", BindingFlags.Static | BindingFlags.NonPublic);
        object[] arguments = { new NoSelectionEtHost(application), null, Activator.CreateInstance(scope, true) };
        IList candidates = (IList)read.Invoke(null, arguments);
        object info = null;
        foreach (object candidate in candidates)
            if (string.Equals((string)candidate.GetType().GetField("WorkbookPath").GetValue(candidate), sourcePath, StringComparison.OrdinalIgnoreCase)) info = candidate;
        Check(info != null, "Application.Selection为空时仍从真实工作簿窗口取得明细：" + arguments[1]);
        IList areas = (IList)info.GetType().GetField("Areas").GetValue(info);
        long count = 0;
        foreach (object area in areas) count += (long)area.GetType().GetProperty("Count").GetValue(area, null);
        Check(areas.Count == 2 && count == 8, "真实窗口RangeSelection保留两片共8格");
        Check((int)areas[0].GetType().GetField("R1").GetValue(areas[0]) == 2
            && (int)areas[1].GetType().GetField("R1").GetValue(areas[1]) == 4, "窗口恢复的来源确为第2行及第4行");
    }

    private static void SetWordCell(ProbeSession session, dynamic doc, object table, int row, int col, string value)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic cell = refs.Take(((dynamic)table).Cell(row, col)), range = refs.Take(cell.Range);
            dynamic text = refs.Take(doc.Range((int)range.Start, (int)range.End - 1)); text.Text = value;
        }
    }

    private static void CheckSelection(ProbeSession session, object application, string expected)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic selection = refs.Take(((dynamic)application).Selection), cells = refs.Take(selection.Cells);
            var coordinates = new List<string>();
            for (int i = 1; i <= (int)cells.Count; i++)
            { dynamic cell = refs.Take(cells[i]); coordinates.Add(cell.RowIndex + "," + cell.ColumnIndex); }
            Check(string.Join(";", coordinates.ToArray()) == expected, "实际选区仅包含指定单元格：" + string.Join(";", coordinates.ToArray()));
        }
    }

    private static void CheckReference(ProbeSession session, object table, int row, int col, string expected)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic cell = refs.Take(((dynamic)table).Cell(row, col)), range = refs.Take(cell.Range), fields = refs.Take(range.Fields);
            Check((int)fields.Count == 1, "每个目标格只有一条LINK");
            dynamic field = refs.Take(fields[1]), code = refs.Take(field.Code);
            List<string> source; string error;
            Check(LinkPasteEngine.TryListSourceRefs((string)code.Text, out source, out error)
                && source.Count == 1 && source[0].EndsWith("!" + expected, StringComparison.Ordinal), "逐格引用正确：" + expected);
        }
    }

    private static string OutsideSignature(ProbeSession session, object table)
    {
        var text = new StringBuilder();
        for (int row = 1; row <= 3; row++)
            for (int col = 1; col <= (row == 1 ? 1 : 4); col++)
            {
                if (row > 1 && col <= 2) continue;
                using (ProbeReferences refs = session.References())
                {
                    dynamic cell = refs.Take(((dynamic)table).Cell(row, col)), range = refs.Take(cell.Range), font = refs.Take(range.Font);
                    text.Append(row).Append(',').Append(col).Append(':').Append((string)range.Text).Append('|')
                        .Append(Convert.ToString(font.Bold)).Append('|').Append(Convert.ToString(font.Size)).Append('|')
                        .Append(Convert.ToString(range.HighlightColorIndex)).Append(';');
                }
            }
        return text.ToString();
    }

    private static string Structure(ProbeSession session, object table)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic range = refs.Take(((dynamic)table).Range);
            var document = new XmlDocument { XmlResolver = null }; document.LoadXml((string)range.WordOpenXML);
            var ns = new XmlNamespaceManager(document.NameTable); ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var text = new StringBuilder();
            foreach (XmlNode node in document.SelectNodes("//w:body/w:tbl/w:tblPr|//w:body/w:tbl/w:tblGrid|//w:body/w:tbl/w:tr/w:tc/w:tcPr", ns)) text.Append(node.OuterXml);
            return text.ToString();
        }
    }

    private static void CheckRejectedUnchanged(ProbeSession session, object application, dynamic doc, LinkPasteStatus expected, string label)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic content = refs.Take(doc.Content), fields = refs.Take(content.Fields), docs = refs.Take(((dynamic)application).Documents);
            string text = (string)content.Text; int count = (int)fields.Count, documents = (int)docs.Count;
            LinkPasteResult result = LinkPasteEngine.Execute(application);
            Check(result.Status == expected && result.CellsLinked == 0, label + "明确拒绝：" + result.Status);
            dynamic after = refs.Take(doc.Content), afterFields = refs.Take(after.Fields);
            Check((string)after.Text == text && (int)afterFields.Count == count && (int)docs.Count == documents,
                label + "未改变文档内容、字段数和文档数");
        }
    }

    private static void PrintLink(string label, LinkPasteResult result)
    {
        Console.WriteLine(label + ": " + result.Status + " linked=" + result.CellsLinked + " refreshed=" + result.Refreshed
            + " target=" + result.TargetRows + "x" + result.TargetCols + " source=" + result.SourceCells + " [" + result.SourceDesc + "]");
        foreach (string failure in result.Failures) Console.WriteLine("  " + failure);
    }

    private static void SetCell(ProbeSession session, object sheet, int row, int col, string value)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic cells = refs.Take((object)((dynamic)sheet).Cells);
            dynamic cell = refs.Take((object)cells[row, col]);
            cell.Value2 = value;
        }
    }

    private static void CopyRange(ProbeSession session, object sheet, string address)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic range = refs.Take((object)((dynamic)sheet).Range(address));
            range.Select();
            range.Copy();
        }
    }

    private static void SelectTable(ProbeSession session, object table)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic range = refs.Take((object)((dynamic)table).Range);
            range.Select();
        }
    }

    private static int Boundary(ProbeSession session, object table, int row, int col, bool end)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic cell = refs.Take((object)((dynamic)table).Cell(row, col));
            dynamic range = refs.Take((object)cell.Range);
            return Convert.ToInt32(end ? range.End : range.Start);
        }
    }

    private static void SelectRect(ProbeSession session, object document, object table, int firstRow, int firstCol, int lastRow, int lastCol)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic range = refs.Take((object)((dynamic)document).Range(
                Boundary(session, table, firstRow, firstCol, false), Boundary(session, table, lastRow, lastCol, true)));
            range.Select();
        }
    }

    private static string CellText(ProbeSession session, object table, int row, int col)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic cell = refs.Take((object)((dynamic)table).Cell(row, col));
            dynamic range = refs.Take((object)cell.Range);
            return (Convert.ToString(range.Text) ?? "").Trim('\r', '\a', '\n', '\0', '\v', '\t', ' ');
        }
    }

    private static void CheckGrid(ProbeSession session, object table, string prefix)
    {
        for (int row = 1; row <= Rows; row++)
            for (int col = 1; col <= Cols; col++)
                Check(CellText(session, table, row, col) == prefix + row + "C" + col, "格(" + row + "," + col + ")=" + prefix + row + "C" + col);
    }

    private static void PreserveDocument(ProbeSession session, object document)
    {
        try
        {
            string path = session.Output("failed-target.docx");
            try { ((dynamic)document).SaveAs2(path, 16); }
            catch { ((dynamic)document).SaveAs(path, 16); }
            Console.WriteLine("失败目标样本：" + path);
        }
        catch (Exception error) { Console.WriteLine("目标样本保存失败：" + error.Message); }
    }

    private static void PrintLogTail()
    {
        try
        {
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordUpdater", "WordUpdater.log");
            if (!File.Exists(path)) return;
            string[] lines = File.ReadAllLines(path);
            for (int i = Math.Max(0, lines.Length - 12); i < lines.Length; i++) Console.WriteLine(lines[i]);
        }
        catch (Exception error) { Console.WriteLine("日志读取失败：" + error.Message); }
    }
}
