using System;
using System.IO;
using WordUpdater;

// 实机测试必须显式运行。Application 可能是用户已有单实例；只关闭自建样本，
// 不退出宿主、不扫描或结束进程。0=PASS，1=FAIL，2=SKIP。
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
