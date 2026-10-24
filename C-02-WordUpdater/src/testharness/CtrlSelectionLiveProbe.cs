using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using WordUpdater;

// Opt-in native regression. Open copies through the normal file association,
// rather than Workbooks.Open on the automation instance being tested. WPS may
// publish an empty automation Application and a separate visible workbook.
// Only copies inside ProbeSession are changed or closed; no Application.Quit.
internal static class CtrlSelectionLiveProbe
{
    [DllImport("ole32.dll")] private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);
    [DllImport("ole32.dll")] private static extern int CreateBindCtx(int reserved, out IBindCtx context);
    private static int checks;

    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        if (args.Length != 2)
        { Console.WriteLine("Usage: CtrlSelectionLiveProbe source.xlsx merged-header-target.docx (only temporary copies are modified)"); return 2; }
        string sourceInput = Path.GetFullPath(args[0]), targetInput = Path.GetFullPath(args[1]);
        string sourceHash = Hash(sourceInput), targetHash = Hash(targetInput);
        using (var session = new ProbeSession("ctrl-selection-instances"))
        {
            object document = null;
            try
            {
                object automation = session.Take(ProbeSession.Connect(new[] { "KET.Application", "Excel.Application" }, false));
                object word = session.Take(ProbeSession.Connect(new[] { "KWPS.Application", "wps.Application" }, false));
                if (automation == null || word == null) return session.Finish(2);
                dynamic hiddenBooks = session.Take(((dynamic)automation).Workbooks);
                if ((int)hiddenBooks.Count != 0)
                { Console.WriteLine("SKIP: needs an existing empty automation instance; no user workbooks were changed."); return session.Finish(2); }
                string sourceCopy = session.CopyInput(sourceInput, "source-own.xlsx");
                using (Process process = Process.Start(new ProcessStartInfo(sourceCopy) {
                    UseShellExecute = true, WindowStyle = ProcessWindowStyle.Hidden })) { }
                object workbook = null;
                var timer = Stopwatch.StartNew();
                while (workbook == null && timer.ElapsedMilliseconds < 10000)
                {
                    workbook = FindOpenWorkbook(session, sourceCopy);
                    if (workbook == null) { Application.DoEvents(); Thread.Sleep(100); }
                }
                if (workbook == null) throw new InvalidOperationException("The normally opened sample did not publish its workbook moniker.");
                session.Own(workbook, () => ((dynamic)workbook).Close(false));
                dynamic book = workbook, excel = session.Take(book.Application), host = word;
                session.Override(() => (object)excel.DisplayAlerts, v => excel.DisplayAlerts = v, false, "source.DisplayAlerts");
                session.Override(() => (object)host.Visible, v => host.Visible = v, true, "Word.Visible");
                session.Override(() => (object)host.DisplayAlerts, v => host.DisplayAlerts = v, 0, "Word.DisplayAlerts");
                dynamic sheets = session.Take(book.Worksheets), sheet = session.Take(sheets[1]), sourceCells = session.Take(sheet.Cells);
                Check((int)hiddenBooks.Count == 0, "the automation instance stays empty while the source is open in another instance");
                dynamic sourceBooks = session.Take(excel.Workbooks);
                Console.WriteLine("OWN SOURCE Workbooks=" + sourceBooks.Count + "; automation Workbooks=" + hiddenBooks.Count);
                dynamic docs = session.Take(host.Documents);
                foreach (int secondRow in new[] { 4, 3 })
                {
                    using (ProbeReferences refs = session.References())
                    {
                        string targetCopy = session.CopyInput(targetInput, "target-rows1-" + secondRow + ".docx");
                        document = docs.Open(targetCopy);
                        object ownDocument = document; session.Own(ownDocument, () => ((dynamic)ownDocument).Close(0));
                        dynamic doc = document, tables = refs.Take(doc.Tables), table = refs.Take(tables[1]);
                        Check(!Convert.ToBoolean(table.Uniform), "sample target retains its merged header");
                        string header = HeaderSignature(refs, table), structure = Structure(refs, table);
                        dynamic source = refs.Take(sheet.Range("A1:C1,A" + secondRow + ":C" + secondRow));
                        book.Activate(); source.Select(); source.Copy();
                        dynamic selection = refs.Take(excel.Selection), areas = refs.Take(selection.Areas);
                        Check((int)areas.Count == 2 && (int)selection.Count == 6, "native source selection contains two areas and six cells");
                        doc.Activate(); SelectTarget(refs, doc, table);
                        CheckTargetSelection(refs, host);
                        Console.WriteLine("READER candidates for own source=" + ReadCandidateCount(sourceCopy));
                        LinkPasteResult linked = LinkPasteEngine.Execute(word);
                        Console.WriteLine("CASE rows1/" + secondRow + ": " + linked.Status + " source=" + linked.SourceCells
                            + " target=" + linked.TargetRows + "x" + linked.TargetCols + " linked=" + linked.CellsLinked + " refreshed=" + linked.Refreshed);
                        Check(linked.Status == LinkPasteStatus.Completed && linked.SourceCells == 6 && linked.CellsLinked == 6
                            && linked.Refreshed == 6 && linked.TargetRow1 == 2 && linked.TargetRows == 2 && linked.TargetCols == 3,
                            "real source instance maps six cells below the merged header: " + linked.ErrorMessage);
                        for (int row = 2; row <= 3; row++)
                            for (int col = 1; col <= 3; col++)
                            {
                                int sourceRow = row == 2 ? 1 : secondRow;
                                dynamic sourceCell = refs.Take(sourceCells[sourceRow, col]), targetCell = refs.Take(table.Cell(row, col));
                                dynamic targetRange = refs.Take(targetCell.Range), fields = refs.Take(targetRange.Fields);
                                Check(((string)targetRange.Text).Trim('\r', '\a') == Convert.ToString(sourceCell.Text), "displayed value matches source R" + sourceRow + "C" + col);
                                Check((int)fields.Count == 1, "one LINK per target cell");
                                dynamic field = refs.Take(fields[1]), code = refs.Take(field.Code);
                                List<string> sources; string error;
                                Check(LinkPasteEngine.TryListSourceRefs((string)code.Text, out sources, out error) && sources.Count == 1
                                    && sources[0].EndsWith("!R" + sourceRow + "C" + col, StringComparison.Ordinal), "LINK refers only to the selected source cell");
                            }
                        dynamic content = refs.Take(doc.Content), allFields = refs.Take(content.Fields);
                        Check((int)allFields.Count == 6, "sample document contains exactly six LINK fields");
                        Check(header == HeaderSignature(refs, table) && structure == Structure(refs, table), "header content, formatting and table geometry are preserved");
                        dynamic first = refs.Take(sourceCells[1, 1]); object original = first.Value2;
                        first.Value2 = "Ctrl update " + secondRow; book.Save(); SelectTarget(refs, doc, table);
                        RunResult updated = UpdateEngine.Execute(word);
                        dynamic firstTarget = refs.Take(table.Cell(2, 1)), firstRange = refs.Take(firstTarget.Range);
                        Check(updated.Status == RunStatus.Completed && updated.Success == 6
                            && ((string)firstRange.Text).Trim('\r', '\a') == "Ctrl update " + secondRow, "six linked cells refresh after changing the copied source");
                        Check(header == HeaderSignature(refs, table) && structure == Structure(refs, table), "refresh also preserves the merged header and geometry");
                        first.Value2 = original; book.Save();
                    }
                }
                Check(Hash(sourceInput) == sourceHash && Hash(targetInput) == targetHash, "both supplied original files are unchanged");
                Console.WriteLine("CtrlSelectionLiveProbe: " + checks + " PASS");
                return session.Finish(0);
            }
            catch (Exception error)
            {
                if (document != null) try { ((dynamic)document).Save(); } catch { }
                session.Fail(error.ToString()); return session.Finish(1);
            }
        }
    }

    private static object FindOpenWorkbook(ProbeSession session, string path)
    {
        using (ProbeReferences refs = session.References())
        {
            IRunningObjectTable rot; Marshal.ThrowExceptionForHR(GetRunningObjectTable(0, out rot)); refs.Take(rot);
            IBindCtx context; Marshal.ThrowExceptionForHR(CreateBindCtx(0, out context)); refs.Take(context);
            IEnumMoniker iterator; rot.EnumRunning(out iterator); refs.Take(iterator);
            IMoniker[] item = new IMoniker[1];
            while (iterator.Next(1, item, IntPtr.Zero) == 0)
            {
                refs.Take(item[0]); string name = null;
                try { item[0].GetDisplayName(context, null, out name); } catch { }
                if (!string.Equals(name, path, StringComparison.OrdinalIgnoreCase)) continue;
                object workbook; rot.GetObject(item[0], out workbook);
                return workbook; // This acquisition is transferred to Own in the caller.
            }
            return null;
        }
    }

    private static int ReadCandidateCount(string path)
    {
        Type reader = typeof(LinkPasteEngine).Assembly.GetType("WordUpdater.EtSelectionReader", true);
        MethodInfo method = reader.GetMethods(BindingFlags.Static | BindingFlags.NonPublic).Single(m => m.Name == "TryReadAll");
        object[] args = method.GetParameters().Length == 1 ? new object[] { null } : new object[] { path, null };
        IList candidates = (IList)method.Invoke(null, args);
        return candidates.Cast<object>().Count(c => string.Equals((string)c.GetType().GetField("WorkbookPath").GetValue(c), path, StringComparison.OrdinalIgnoreCase));
    }

    private static void SelectTarget(ProbeReferences refs, dynamic doc, dynamic table)
    {
        dynamic first = refs.Take(table.Cell(2, 1)), last = refs.Take(table.Cell(3, 3));
        dynamic start = refs.Take(first.Range), end = refs.Take(last.Range), range = refs.Take(doc.Range((int)start.Start, (int)end.End));
        range.Select();
    }

    private static void CheckTargetSelection(ProbeReferences refs, dynamic host)
    {
        dynamic selected = refs.Take(host.Selection), cells = refs.Take(selected.Cells);
        var coordinates = new List<string>();
        for (int i = 1; i <= (int)cells.Count; i++)
        { dynamic cell = refs.Take(cells[i]); coordinates.Add(cell.RowIndex + "," + cell.ColumnIndex); }
        Check(string.Join(";", coordinates.ToArray()) == "2,1;2,2;2,3;3,1;3,2;3,3", "native Word selection is exactly rows2/3, columns1..3");
    }

    private static string HeaderSignature(ProbeReferences refs, dynamic table)
    {
        dynamic header = refs.Take(table.Cell(1, 1)), range = refs.Take(header.Range), font = refs.Take(range.Font), paragraphs = refs.Take(range.ParagraphFormat);
        return range.Text + "|" + font.Name + "|" + font.Size + "|" + font.Bold + "|" + range.HighlightColorIndex + "|" + paragraphs.Alignment;
    }

    private static string Structure(ProbeReferences refs, dynamic table)
    {
        dynamic range = refs.Take(table.Range);
        var xml = new XmlDocument { XmlResolver = null }; xml.LoadXml((string)range.WordOpenXML);
        var ns = new XmlNamespaceManager(xml.NameTable); ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var result = new StringBuilder();
        foreach (XmlNode node in xml.SelectNodes("//w:body/w:tbl/w:tblPr|//w:body/w:tbl/w:tblGrid|//w:body/w:tbl/w:tr/w:tc/w:tcPr", ns)) result.Append(node.OuterXml);
        return result.ToString();
    }

    private static string Hash(string path)
    { using (SHA256 sha = SHA256.Create()) using (FileStream file = File.OpenRead(path)) return BitConverter.ToString(sha.ComputeHash(file)); }
    private static void Check(bool value, string message)
    { if (!value) throw new InvalidOperationException(message); checks++; Console.WriteLine("PASS: " + message); }
}
