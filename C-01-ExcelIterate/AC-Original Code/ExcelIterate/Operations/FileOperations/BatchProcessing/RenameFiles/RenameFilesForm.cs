using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class RenameWorkbookSession
    {
        internal static readonly string[] Headers = { "文档路径", "文件现有名称", "新文件名（待填）", "处理状态" };
        private readonly object app; private readonly string bookPath;
        private string anchor, folder; private List<FileSnapshot> files;
        internal RenameWorkbookSession(object application)
        {
            app = application;
            using (var bag = new ComBag())
            {
                object book = bag.Own(((dynamic)app).ActiveWorkbook);
                if (book == null) throw new InvalidOperationException("请先打开用于接收重命名清单的工作簿。");
                Connect.PrepareWorkbookSave(Convert.ToString(((dynamic)book).FullName));
                bookPath = WorkbookSaving.Save(book, false);
            }
        }
        private object Book(ComBag bag)
        {
            object book = bag.Own(LinkAssistantSession.FindBook(app, bookPath));
            if (book == null) throw new InvalidOperationException("打开窗口时绑定的清单工作簿已关闭，请重新打开窗口。");
            return book;
        }
        private object Sheet(object book, ComBag bag)
        {
            if (anchor == null) throw new InvalidOperationException("请先生成文件清单。");
            dynamic names = bag.Own(((dynamic)book).Names), name = bag.Own(names.Item(anchor)), range = bag.Own(name.RefersToRange);
            return bag.Own(range.Worksheet);
        }
        internal string Create(string root, List<FileSnapshot> scanned)
        {
            using (var bag = new ComBag())
            {
                dynamic book = Book(bag); WorkbookSaving.Save((object)book, false);
                using (var quiet = new HostQuiet(app))
                {
                    object sheet = WorkbookUi.EnsureSheet((object)book, "文件批量命名", bag, false);
                    string token = "_EI_Rename_" + Guid.NewGuid().ToString("N");
                    dynamic names = bag.Own(book.Names); bag.Own(names.Add(token, "=" + WorkbookUi.Quote(Convert.ToString(((dynamic)sheet).Name)) + "!$A$1", false));
                    object[,] values = new object[scanned.Count + 1, 4];
                    for (int c = 0; c < 4; c++) values[0,c] = Headers[c];
                    for (int i = 0; i < scanned.Count; i++) { values[i+1,0] = scanned[i].Path; values[i+1,1] = scanned[i].Name; values[i+1,2] = ""; values[i+1,3] = "待填写"; }
                    dynamic range = bag.Own(((dynamic)sheet).Range["A1:D" + (scanned.Count+1)]); range.NumberFormat = "@"; range.Value2 = values;
                    WorkbookUi.StyleTable(sheet, 4, scanned.Count + 1);
                    WorkbookUi.Center(range);
                    WorkbookUi.FitOutput(sheet, 1, 4, scanned.Count + 1);
                    ((dynamic)sheet).Activate(); folder = root; files = scanned; anchor = token;
                    WorkbookSaving.Save((object)book, false);
                    return "已在绑定工作簿中生成「" + Convert.ToString(((dynamic)sheet).Name) + "」，共 " + files.Count + " 个文件。填写第三列后点击执行替换。";
                }
            }
        }
        internal RenamePlan Read()
        {
            using (var bag = new ComBag())
            {
                object book = Book(bag); Connect.PrepareWorkbookSave(bookPath); WorkbookSaving.Save(book, false);
                dynamic sheet = Sheet(book, bag), used = bag.Own(sheet.UsedRange), rows = bag.Own(used.Rows);
                if (Convert.ToInt32(rows.Count) != files.Count + 1) throw new InvalidOperationException("清单行数已变化，请重新生成清单；不要增删行。");
                dynamic range = bag.Own(sheet.Range["A1:D" + (files.Count+1)]); Array data = range.Value2 as Array;
                var plan=Parse(folder, files, data);plan.DocumentPath=bookPath;plan.Anchor=anchor;return plan;
            }
        }
        internal static RenamePlan Parse(string root, IList<FileSnapshot> knownFiles, Array data)
        {
            if (data == null || data.Rank != 2 || data.GetLength(0) != knownFiles.Count+1 || data.GetLength(1) != 4) throw new InvalidOperationException("无法读取完整清单，或行数已变化。");
            int r0 = data.GetLowerBound(0), c0 = data.GetLowerBound(1);
            for (int c = 0; c < 4; c++) if (Convert.ToString(data.GetValue(r0,c0+c)) != Headers[c]) throw new InvalidOperationException("清单表头已变化，请重新生成清单。");
            var known = knownFiles.ToDictionary(f => f.Path, StringComparer.OrdinalIgnoreCase); var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase); var plan = new RenamePlan { Root = root };
            for (int r = 1; r <= knownFiles.Count; r++)
            {
                string path = Convert.ToString(data.GetValue(r0+r,c0)); FileSnapshot file;
                if (!known.TryGetValue(path, out file) || !seen.Add(path) || path != file.Path || Convert.ToString(data.GetValue(r0+r,c0+1)) != file.Name) throw new InvalidOperationException("原路径或原文件名已变化，或有重复行；请重新生成清单。");
                plan.Items.Add(new RenameItem { Source = file, Input = Convert.ToString(data.GetValue(r0+r,c0+2)) });
            }
            return plan;
        }
        internal void Restore(OperationBatch batch)
        {
            if(batch.Kind!="Rename" || !string.Equals(batch.DocumentPath,bookPath,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("请打开原清单工作簿后恢复该批次。");
            folder=batch.Root;anchor=batch.Anchor;files=batch.ReportFiles.Select(f=>f.Snapshot()).ToList();
        }
        internal void WriteResult(FileOperationResult result)
        {
            using (var bag = new ComBag())
            {
                object book = Book(bag); dynamic sheet = Sheet(book, bag), used = bag.Own(sheet.UsedRange), rows = bag.Own(used.Rows);
                if (Convert.ToInt32(rows.Count) != files.Count+1) throw new InvalidOperationException("处理期间清单行数已变化，文件结果请以日志为准。");
                dynamic range = bag.Own(sheet.Range["A1:D" + (files.Count+1)]);
                Array data = range.Value2 as Array; var plan = Parse(folder, files, data);
                for (int i = 0; i < plan.Items.Count; i++)
                {
                    string status; if (!result.Status.TryGetValue(plan.Items[i].Source.Path, out status)) status = result.Failed > 0 ? "整批暂停，未执行" : "未执行";
                    dynamic cell = bag.Own(sheet.Range["D" + (i+2)]); cell.Value2 = status;
                }
                WorkbookUi.FitOutput((object)sheet, 1, 4, files.Count + 1);
                sheet.Activate(); WorkbookSaving.Save(book, false);
            }
        }
    }

    internal sealed class RenameFilesForm : FileToolForm
    {
        private readonly RenameWorkbookSession session;
        internal readonly Button Scan=Ui.Button("生成文件清单",140),Execute=Ui.Button("执行替换",112);
        internal RenameFilesForm(RenameWorkbookSession bound):base("文件批量命名","file.rename")
        {
            session=bound;Introduction="1. 请浏览选择文件夹，或自行粘贴带引号的完整路径。\r\n2. 只读取第一层文件；第三列可只填文件名或完整文件名（文件名+扩展名）。";SetLog("");Status.Text="";
            Actions.Controls.Add(Scan);Actions.Controls.SetChildIndex(Scan,0);Actions.Controls.Add(Execute);Actions.Controls.SetChildIndex(Execute,1);
            Scan.Click+=delegate {if(!Ready())return;string root;try{root=FilePaths.Folder(FolderPath.Text);}catch(Exception ex){Status.Text=ex.Message;return;}
                Run(delegate {var notes=new List<string>();return FilePaths.Scan(root,false,notes);},delegate(List<FileSnapshot> files){if(CancelRequested){Status.Text="已取消，未生成清单。";return;}if(session==null){Status.Text="预览窗口不连接工作簿。";return;}Status.Text=session.Create(root,files);SetLog(Status.Text);});};
            Execute.Click+=delegate {if(!Ready())return;RenamePlan plan;try {if(session==null)throw new InvalidOperationException("预览窗口不连接工作簿。");plan=session.Read();}catch(Exception ex){Status.Text=ex.Message;LogView.AppendText("\r\n"+ex.Message);return;}
                Run(()=>FileRenamer.Execute(plan,()=>CancelRequested,Progress,null),delegate(FileOperationResult result){WriteStatus(result);Result(result);});};
        }
        protected override void RestoreBatch(OperationBatch batch){if(session!=null)session.Restore(batch);}
        protected override void UndoFinished(FileOperationResult result){WriteStatus(result);Result(result);}
        private void WriteStatus(FileOperationResult result)
        {try {if(session!=null)session.WriteResult(result);result.Log.AppendLine("清单状态已回写并自动保存。");}catch(Exception ex){result.Log.AppendLine("清单回写失败，文件实际结果仍有效："+ex.Message);}if(result.Batch!=null)result.Batch.Log(result.Text);}
    }
}
