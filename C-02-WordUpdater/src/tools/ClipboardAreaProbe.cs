using System;
using System.IO;
using System.Text;
using WordUpdater;

// Live COM observations, not a Ribbon or mouse/keyboard acceptance test.
// Only explicitly created probe documents are closed. An Application obtained
// through CreateInstance can belong to the user and is never Quit or killed.
internal static class ClipboardAreaProbe
{
    private static readonly string[] ExcelIds = { "KET.Application", "et.Application", "Excel.Application" };
    private static readonly string[] WordIds = { "KWPS.Application", "wps.Application", "Word.Application" };

    private static void SetCell(ProbeSession session, object worksheet, int row, int col, string value)
    {
        using (var refs = session.References())
        {
            dynamic ws = worksheet;
            dynamic cells = refs.Take(ws.Cells);
            dynamic cell = refs.Take(cells[row, col]);
            cell.Value2 = value;
        }
    }

    private static int CellBoundary(ProbeSession session, object table, int row, int col, bool end)
    {
        using (var refs = session.References())
        {
            dynamic cell = refs.Take(((dynamic)table).Cell(row, col));
            dynamic range = refs.Take(cell.Range);
            return Convert.ToInt32(end ? range.End : range.Start);
        }
    }

    private static TargetRectStatus ProbeRect(ProbeSession session, string label, object application, object table)
    {
        using (var refs = session.References())
            return ReadRect(session, label, refs.Take(((dynamic)application).Selection), table);
    }

    internal static TargetRectStatus ReadRect(ProbeSession session, string label, object selection, object table)
    {
        using (var refs = session.References())
        {
            dynamic sel = selection;
            LinkTargetRect rect; string detail;
            TargetRectStatus status = LinkPasteGeometry.ResolveTargetRect(selection, table, out rect, out detail);
            dynamic range = refs.Take(sel.Range);
            dynamic tableRange = refs.Take(((dynamic)table).Range);
            dynamic cells = null;
            int count = -1;
            try { cells = refs.Take(sel.Cells); if (cells != null) count = Convert.ToInt32(cells.Count); }
            catch (Exception error) { Console.WriteLine("  Selection.Cells unavailable: " + error.Message); }
            if (count < 0)
            {
                try { cells = refs.Take(range.Cells); if (cells != null) count = Convert.ToInt32(cells.Count); }
                catch (Exception error) { Console.WriteLine("  Range.Cells unavailable: " + error.Message); }
            }
            if (status == TargetRectStatus.Ok && count < 0)
                throw new InvalidOperationException("Valid rectangle has no readable diagnostic cells.");
            Console.WriteLine(label + ": Cells=" + (count < 0 ? "unavailable" : count.ToString()) + " Selection=[" + range.Start + "," + range.End
                + "] Table=[" + tableRange.Start + "," + tableRange.End + "]");
            string positions = "";
            for (int i = 1; i <= count && i <= 12; i++)
            {
                dynamic cell = refs.Take(cells.Item(i));
                positions += "(" + cell.RowIndex + "," + cell.ColumnIndex + ")";
            }
            Console.WriteLine("  positions=" + positions);
            Console.WriteLine("  ResolveTargetRect=" + status + " " + (rect == null ? "" : rect.ToString()) + " " + detail);
            return status;
        }
    }

    internal static void RequireStatus(TargetRectStatus actual, params TargetRectStatus[] expected)
    {
        foreach (TargetRectStatus value in expected) if (actual == value) return;
        throw new InvalidOperationException("Unexpected target geometry: " + actual);
    }

    // Insert the second table into an explicit interior paragraph while the new
    // document still has only text. Adding the first table then shifts this owned
    // second-table Range automatically; no guessed final-paragraph position is used.
    internal static object[] CreateBoundaryTables(ProbeSession session, object document)
    {
        const string outside = "CLIPBOARD_PROBE_OUTSIDE";
        const string anchor = "CLIPBOARD_PROBE_SECOND_TABLE_ANCHOR";
        string paragraphs = "\r" + outside + "\r" + anchor + "\r";
        using (var refs = session.References())
        {
            dynamic doc = document;
            dynamic content = refs.Take(doc.Content);
            content.Text = paragraphs;
            int beginning = Convert.ToInt32(content.Start);
            int position = beginning + paragraphs.IndexOf(anchor, StringComparison.Ordinal);
            dynamic anchorRange = refs.Take(doc.Range(position, position + anchor.Length));
            if (Convert.ToString(anchorRange.Text) != anchor)
                throw new InvalidOperationException("Second-table text anchor was not found at the verified document offset.");
            dynamic tables = refs.Take(doc.Tables);
            dynamic insertion = refs.Take(doc.Range(position, position));
            object second = session.Take((object)tables.Add(insertion, 2, 2));
            dynamic start = refs.Take(doc.Range(beginning, beginning));
            object first = session.Take((object)tables.Add(start, 3, 3));
            dynamic firstRange = refs.Take(((dynamic)first).Range);
            dynamic secondRange = refs.Take(((dynamic)second).Range);
            int firstEnd = Convert.ToInt32(firstRange.End), secondStart = Convert.ToInt32(secondRange.Start);
            if (Convert.ToInt32(tables.Count) != 2 || secondStart <= firstEnd)
                throw new InvalidOperationException("Boundary sample must contain two separate, ordered tables.");
            dynamic gap = refs.Take(doc.Range(firstEnd, secondStart));
            if (!(Convert.ToString(gap.Text) ?? "").Contains(outside))
                throw new InvalidOperationException("Boundary sample has no verified outside-table paragraph.");
            Console.WriteLine("BOUNDARY_SAMPLE: table1=[" + firstRange.Start + "," + firstEnd + "] outside paragraph; table2=["
                + secondStart + "," + secondRange.End + "]");
            return new[] { first, second };
        }
    }

    internal static void ProbeRejection(ProbeSession session, string label, object application, object document, object table)
    {
        TargetRectStatus geometry = ProbeRect(session, label, application, table);
        RequireStatus(geometry, TargetRectStatus.OutOfTable, TargetRectStatus.MultiTable);
        using (var refs = session.References())
        {
            dynamic doc = document;
            dynamic documents = refs.Take(((dynamic)application).Documents);
            dynamic content = refs.Take(doc.Content);
            dynamic fields = refs.Take(content.Fields);
            string beforeText = Convert.ToString(content.Text);
            int beforeFields = Convert.ToInt32(fields.Count), beforeDocuments = Convert.ToInt32(documents.Count);
            LinkPasteResult result = LinkPasteEngine.Execute(application);
            if ((result.Status != LinkPasteStatus.NotInTable && result.Status != LinkPasteStatus.SelectionOutOfTable) || result.CellsLinked != 0)
                throw new InvalidOperationException("Negative selection was not explicitly rejected: " + result.Status + "; " + result.ErrorMessage);
            dynamic after = refs.Take(doc.Content);
            dynamic afterFields = refs.Take(after.Fields);
            if (Convert.ToString(after.Text) != beforeText || Convert.ToInt32(afterFields.Count) != beforeFields
                || Convert.ToInt32(documents.Count) != beforeDocuments)
                throw new InvalidOperationException("Rejecting a negative selection changed document content, fields or document count.");
            Console.WriteLine("PASS: " + label + " engine=" + result.Status + "; 0 fields written; content/field/document counts unchanged");
        }
    }

    private static string PasteAndReadLink(ProbeSession session, object application, int dataType)
    {
        object document = null;
        using (var refs = session.References())
        {
            try
            {
                dynamic documents = refs.Take(((dynamic)application).Documents);
                dynamic doc = refs.Take(documents.Add()); document = doc;
                dynamic selection = refs.Take(((dynamic)application).Selection);
                try { selection.PasteSpecial(0, true, 0, false, dataType); }
                catch (Exception error) { return "PasteSpecial unavailable: " + error.Message; }
                dynamic content = refs.Take(doc.Content);
                dynamic fields = refs.Take(content.Fields);
                if (Convert.ToInt32(fields.Count) > 0)
                {
                    dynamic field = refs.Take(fields.Item(1));
                    dynamic code = refs.Take(field.Code);
                    return Convert.ToString(code.Text);
                }
                dynamic shapes = refs.Take(content.InlineShapes);
                if (Convert.ToInt32(shapes.Count) > 0)
                {
                    dynamic shape = refs.Take(shapes.Item(1));
                    dynamic field = refs.Take(shape.Field);
                    dynamic code = refs.Take(field.Code);
                    return Convert.ToString(code.Text);
                }
                return "No linked field in this paste format.";
            }
            finally
            {
                if (document != null)
                    try { ((dynamic)document).Close(0); }
                    catch (Exception error) { session.Fail("closing temporary paste document: " + error.Message); }
            }
        }
    }

    private static void Observe(ProbeSession session, string label, Action action)
    {
        try { action(); }
        catch (Exception error) { session.Fail(label + ": " + error.Message); }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        try { return Run(); }
        catch (Exception error) { Console.WriteLine("FAIL: " + error); return 1; }
    }

    private static int Run()
    {
        using (var session = new ProbeSession("clipboard-area"))
        {
            int result = 1;
            object workbook = null, document = null;
            try
            {
                object et = session.Take(ProbeSession.Connect(ExcelIds, false));
                object wps = session.Take(ProbeSession.Connect(WordIds, false));
                if (et == null || wps == null)
                {
                    Console.WriteLine("SKIP: WPS COM host is unavailable; no acceptance result.");
                    result = 2;
                }
                else
                {
                    dynamic excel = et, word = wps;
                    session.Override(() => (object)excel.DisplayAlerts, value => excel.DisplayAlerts = value, false, "Excel.DisplayAlerts");
                    session.Override(() => (object)word.DisplayAlerts, value => word.DisplayAlerts = value, 0, "Word.DisplayAlerts");
                    string source = session.Output("source.xlsx");
                    dynamic books = session.Take(excel.Workbooks);
                    workbook = session.Own((object)books.Add(), () => ((dynamic)workbook).Close(false));
                    dynamic book = workbook;
                    dynamic sheets = session.Take(book.Worksheets);
                    dynamic sheet = session.Take(sheets.Item(1));
                    for (int row = 1; row <= 3; row++)
                        for (int col = 1; col <= 8; col++) SetCell(session, sheet, row, col, "R" + row + "C" + col);
                    book.SaveAs(source);

                    dynamic documents = session.Take(word.Documents);
                    document = session.Own((object)documents.Add(), () => ((dynamic)document).Close(0));
                    dynamic doc = document;
                    object[] boundaryTables = CreateBoundaryTables(session, document);
                    object table = boundaryTables[0], second = boundaryTables[1];

                    foreach (string address in new[] { "H1,H3", "H1:H3" })
                    {
                        string copyAddress = address;
                        Observe(session, "copy " + address, delegate {
                            dynamic range = session.Take(sheet.Range(copyAddress)); range.Select(); range.Copy();
                            dynamic selection = session.Take(excel.Selection);
                            dynamic areas = session.Take(selection.Areas);
                            Console.WriteLine(copyAddress + " Selection.Areas=" + areas.Count);
                            for (int i = 1; i <= Convert.ToInt32(areas.Count); i++)
                            {
                                using (var refs = session.References())
                                {
                                    dynamic area = refs.Take(areas.Item(i));
                                    dynamic ws = refs.Take(area.Worksheet);
                                    Console.WriteLine("  " + area.Address + " Sheet=" + ws.Name);
                                }
                            }
                            foreach (int type in new[] { 2, 1, 0 })
                                Console.WriteLine("  Paste type " + type + ": " + PasteAndReadLink(session, wps, type));
                        });
                    }

                    Observe(session, "clipboard formats", delegate {
                        var data = System.Windows.Forms.Clipboard.GetDataObject();
                        if (data == null) throw new InvalidOperationException("Clipboard has no data object.");
                        foreach (string format in data.GetFormats(false)) Console.WriteLine("Clipboard: " + format);
                    });
                    Observe(session, "Q4 insertion point", delegate {
                        dynamic cell = session.Take(((dynamic)table).Cell(2, 2));
                        dynamic range = session.Take(cell.Range); range.Collapse(1); range.Select();
                        RequireStatus(ProbeRect(session, "Q4 insertion point", wps, table), TargetRectStatus.Ok);
                    });
                    int[,] bounds = { { 3, 3, 3, 3 }, { 1, 1, 2, 2 }, { 1, 1, 1, 3 } };
                    for (int index = 0; index < bounds.GetLength(0); index++)
                    {
                        int i = index;
                        Observe(session, "Q5-" + i, delegate {
                            dynamic range = session.Take(doc.Range(
                                CellBoundary(session, table, bounds[i, 0], bounds[i, 1], false),
                                CellBoundary(session, table, bounds[i, 2], bounds[i, 3], true)));
                            range.Select(); RequireStatus(ProbeRect(session, "Q5-" + i, wps, table), TargetRectStatus.Ok);
                        });
                    }
                    Observe(session, "Q5 whole table", delegate {
                        dynamic range = session.Take(((dynamic)table).Range); range.Select();
                        RequireStatus(ProbeRect(session, "Q5 whole table", wps, table), TargetRectStatus.Ok);
                    });
                    Observe(session, "Q6 outside table", delegate {
                        dynamic secondRange = session.Take(((dynamic)second).Range);
                        dynamic range = session.Take(doc.Range(CellBoundary(session, table, 3, 3, false), Convert.ToInt32(secondRange.Start) - 1));
                        range.Select(); ProbeRejection(session, "Q6 outside table", wps, document, table);
                    });
                    Observe(session, "Q7 cross tables", delegate {
                        dynamic range = session.Take(doc.Range(CellBoundary(session, table, 3, 3, false), CellBoundary(session, second, 1, 1, true)));
                        range.Select(); ProbeRejection(session, "Q7 cross tables", wps, document, table);
                    });
                    result = 0;
                }
            }
            catch (Exception error) { session.Fail(error.ToString()); result = 1; }
            return session.Finish(result);
        }
    }
}
