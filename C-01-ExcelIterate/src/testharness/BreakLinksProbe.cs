using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Excel_Iterate;

partial class WorkbookToolsProbe
{
    static void BreakLinkChecks(string folder,string a,string b,string c)
    {
        string target=Path.Combine(folder,"断链 #100% .xlsx");
        Package(target,a,b);WriteFormulaSample(target,a,b);
        byte[] before=File.ReadAllBytes(target);
        var session=JournalSession(c,new[]{target},a,"  "+LinkAssistantSession.BreakCommand+"  ",true);
        var book=((LinkFakeApp)session.App).Workbooks.Items[1];
        session.Execute(null);
        Check(session.Batch.Changes.Single().Action==LinkAction.Break&&session.Batch.Changes[0].Applied,"exact trimmed command persists a break action");
        Check(XlsxLinkPackage.Scan(target).Select(l=>l.SourcePath).SequenceEqual(new[]{b}),"break removes selected source and preserves other link");
        var xml=SheetXml(target);Check(xml.SelectSingleNode("//c[@r='A1']/f")==null&&xml.SelectSingleNode("//c[@r='A1']/v").InnerText=="12345","break converts selected formula to cached value");
        Check(xml.SelectSingleNode("//c[@r='B1']/f")!=null&&xml.SelectSingleNode("//c[@r='B1']/v").InnerText=="10","unselected formula and value preserved");
        var checkpoint=session.Batch.Checkpoints.Single();Check(File.ReadAllBytes(checkpoint.Backup).SequenceEqual(before)&&checkpoint.LastHash==XlsxLinkPackage.Hash(target),"checkpoint preserves full original bytes and committed hash");
        foreach(string step in session.Steps(true,()=>false)){}
        Check(session.Batch.State=="UndoFailed"&&book.CloseCount==0&&WpsShared.SourcePaths.Exists(checkpoint.Backup),"undo never closes an existing workbook and retains backup");Reject(()=>session.Batch.Confirm(_=>{}),"pending whole-book restore blocks confirm");
        book.Close(false);var reopened=new LinkAssistantSession(session.App);foreach(string step in reopened.Steps(true,()=>false)){}
        Check(File.ReadAllBytes(target).SequenceEqual(before)&&!reopened.Batch.CanUndo,"reopened journal restores all original formulas byte for byte");reopened.Batch.Confirm(File.Delete);Check(!WpsShared.SourcePaths.Exists(checkpoint.Backup),"confirm removes only recorded verified checkpoint");

        Package(target,a,b);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        string result=session.Execute(null);Check(result.Contains("成功1条")&&XlsxLinkPackage.Scan(target).Count==1,"closed workbook uses existing host stage and commits selected break");Check(((LinkFakeApp)session.App).Workbooks.Count==1,"closed stage workbook is the only one closed");
        byte[] submitted=File.ReadAllBytes(target);using(var stream=File.Open(target,FileMode.Open,FileAccess.ReadWrite))using(var zip=new ZipArchive(stream,ZipArchiveMode.Update)){zip.GetEntry("customXml/item1.xml").Delete();Part(zip,"customXml/item1.xml","<content>new user edit</content>");}
        string edited=XlsxLinkPackage.Hash(target);foreach(string step in session.Steps(true,()=>false)){}
        Check(session.Batch.State=="UndoFailed"&&XlsxLinkPackage.Hash(target)==edited&&WpsShared.SourcePaths.Exists(session.Batch.Checkpoints[0].Backup),"post-break edits block whole-book overwrite");
        File.WriteAllBytes(target,submitted);foreach(string step in session.Steps(true,()=>false)){}Check(File.ReadAllBytes(target).SequenceEqual(before),"unchanged submitted bytes permit retry restore");session.Batch.Confirm(File.Delete);

        Package(target,a,b);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        session.Links.Add(new ExternalLinkInfo {OwnerPath=target,SourcePath=b});var report=((LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook).Sheet;
        report.Values=new object[,]{{"文档路径","外部链接","新链接（待填）","处理状态"},{target,a,LinkAssistantSession.BreakCommand,""},{target,b,c,""}};
        session.Execute(null);Check(XlsxLinkPackage.Scan(target).Select(l=>l.SourcePath).SequenceEqual(new[]{c})&&session.Batch.Changes.All(x=>x.Applied),"mixed owner replaces first and breaks only specified source");
        foreach(string step in session.Steps(true,()=>false)){}Check(File.ReadAllBytes(target).SequenceEqual(before),"whole-book undo restores both break and same-owner replacement");session.Batch.Confirm(File.Delete);

        Package(target,a,b);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        session.Links.Add(new ExternalLinkInfo {OwnerPath=target,SourcePath=b});report=((LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook).Sheet;
        report.Values=new object[,]{{"文档路径","外部链接","新链接（待填）","处理状态"},{target,a,LinkAssistantSession.BreakCommand,""},{target,b,a,""}};
        session.Execute(null);Check(File.ReadAllBytes(target).SequenceEqual(before)&&session.Batch.Checkpoints.Count==0,"replacement cannot merge into a source being broken");session.Batch.Confirm(File.Delete);

        foreach(bool refuse in new[]{true,false})
        {
            Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,true);book=((LinkFakeApp)session.App).Workbooks.Items[1];book.RefuseBreak=refuse;book.NoBreak=!refuse;
            session.Execute(null);Check(session.Batch.State=="UndoFailed"&&book.CloseCount==0,"unsupported or ineffective native break never reported as complete");book.Close(false);foreach(string step in session.Steps(true,()=>false)){}Check(File.ReadAllBytes(target).SequenceEqual(before),"failure recovery retains original formula bytes");session.Batch.Confirm(File.Delete);
        }
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,true);book=((LinkFakeApp)session.App).Workbooks.Items[1];book.FailSaveAfterBreak=true;
        session.Execute(null);Check(session.Batch.State=="UndoFailed"&&WpsShared.SourcePaths.Exists(session.Batch.Checkpoints[0].Backup),"save failure leaves recoverable intent and backup");book.Close(false);foreach(string step in session.Steps(true,()=>false)){}Check(File.ReadAllBytes(target).SequenceEqual(before),"save failure can recover after owner closed");session.Batch.Confirm(File.Delete);

        string missing=Path.Combine(folder,"already missing.xlsx");Package(target,missing);session=JournalSession(c,new[]{target},missing,LinkAssistantSession.BreakCommand,true);session.Execute(null);Check(XlsxLinkPackage.Scan(target).Count==0,"missing source file still supports break from cached values");book=((LinkFakeApp)session.App).Workbooks.Items[1];book.Close(false);foreach(string step in session.Steps(true,()=>false)){}session.Batch.Confirm(File.Delete);
        Package(target,a);session=JournalSession(c,new[]{target},a,"请断开该链接！",true);Reject(()=>session.Execute(null),"near command text is not a break action");Check(!session.Links[0].BreakRequested,"near command leaves action unset");
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        bool stop=false;var iterator=session.Steps(false,()=>stop).GetEnumerator();Check(iterator.MoveNext(),"cancellation fixture reaches checkpoint");stop=true;while(iterator.MoveNext()){}iterator.Dispose();Check(File.ReadAllBytes(target).SequenceEqual(before)&&!session.Batch.CanUndo,"cancellation before host call commits nothing");session.Batch.Confirm(File.Delete);
        Package(target,a);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,true);book=((LinkFakeApp)session.App).Workbooks.Items[1];book.SourceStyle="uri";book.RetainHistorical=true;session.Execute(null);Check(XlsxLinkPackage.Scan(target).Count==0,"raw URI name and stale in-memory source handled");book.Close(false);foreach(string step in session.Steps(true,()=>false)){}session.Batch.Confirm(File.Delete);
        var legacy="<LinkBatch><ReportBook>"+System.Security.SecurityElement.Escape(c)+"</ReportBook><Changes><LinkChange><Owner>old</Owner><Old>A</Old><New>B</New></LinkChange></Changes></LinkBatch>";
        using(var reader=new StringReader(legacy)){var old=(LinkBatch)new XmlSerializer(typeof(LinkBatch)).Deserialize(reader);Check(old.Changes[0].Action==LinkAction.Replace&&old.Checkpoints.Count==0,"legacy journals default to replace and no checkpoint");}
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,true);book=((LinkFakeApp)session.App).Workbooks.Items[1];
        iterator=session.Steps(false,()=>false).GetEnumerator();Check(iterator.MoveNext(),"new edit fixture reaches backup");book.Saved=false;while(iterator.MoveNext()){}iterator.Dispose();Check(!book.Saved&&book.SaveCount==0&&File.ReadAllBytes(target).SequenceEqual(before),"new unsaved edits after checkpoint are never silently saved or overwritten");session.Batch.Confirm(File.Delete);
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        stop=false;iterator=session.Steps(false,()=>stop).GetEnumerator();Check(iterator.MoveNext()&&iterator.MoveNext(),"stage cancel fixture reaches native stage result");Check(File.ReadAllBytes(target).SequenceEqual(before),"stage native call leaves original untouched until commit");stop=true;while(iterator.MoveNext()){}iterator.Dispose();Check(File.ReadAllBytes(target).SequenceEqual(before)&&((LinkFakeApp)session.App).Workbooks.Count==1,"stage cancellation closes owned copy and restores nothing into original");session.Batch.Confirm(File.Delete);
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;session.Execute(null);
        checkpoint=session.Batch.Checkpoints[0];checkpoint.PreparedHash=checkpoint.LastHash;checkpoint.LastHash=checkpoint.BeforeHash;checkpoint.CommitIntent=true;checkpoint.Applied=false;session.Batch.Changes[0].Applied=false;session.Batch.Changes[0].Intent=true;session.Batch.Persist();
        reopened=new LinkAssistantSession(session.App);foreach(string step in reopened.Steps(true,()=>false)){}Check(File.ReadAllBytes(target).SequenceEqual(before)&&!reopened.Batch.CanUndo,"commit interruption is reconciled by prepared digest then restored");reopened.Batch.Confirm(File.Delete);
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;session.Execute(null);checkpoint=session.Batch.Checkpoints[0];
        byte[] backupBytes=File.ReadAllBytes(checkpoint.Backup);File.AppendAllText(checkpoint.Backup,"damaged");string appliedHash=XlsxLinkPackage.Hash(target);foreach(string step in session.Steps(true,()=>false)){}Check(session.Batch.State=="UndoFailed"&&XlsxLinkPackage.Hash(target)==appliedHash,"damaged backup blocks restore and leaves current data untouched");File.WriteAllBytes(checkpoint.Backup,backupBytes);foreach(string step in session.Steps(true,()=>false)){}session.Batch.Confirm(File.Delete);
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(target,new[]{target},a,LinkAssistantSession.BreakCommand,false);session.Execute(null);Check(session.Batch.Checkpoints[0].LastHash==XlsxLinkPackage.Hash(target),"receiver checkpoint includes program report-write digest");
        var fittedReceiver=(LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook;
        Check(fittedReceiver.SavedReports.Last()==fittedReceiver.Sheet.State()&&fittedReceiver.SavedReports.Last().Contains("已断开"),"break receiver saves final status together with the complete fitted formats");
        using(var stream=File.OpenRead(target))using(var zip=new ZipArchive(stream,ZipArchiveMode.Read))using(var reader=new StreamReader(zip.GetEntry("customXml/probe-report.xml").Open()))Check(reader.ReadToEnd()==fittedReceiver.SavedReports.Last(),"break checkpoint fingerprints actual disk report values and formats, not only a Save call");
        Check(session.Batch.Checkpoints[0].LastHash==fittedReceiver.SavedHashes.Last()&&fittedReceiver.Sheet.Cell(1,4).Bold&&fittedReceiver.Sheet.Cell(1,4).Horizontal==-4108&&!fittedReceiver.Sheet.Cell(1,4).Wrap,"break checkpoint is recorded after final bold centered single-line formatting reaches disk");
        ((LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook).Close(false);foreach(string step in session.Steps(true,()=>false)){}Check(File.ReadAllBytes(target).SequenceEqual(before),"receiver can restore while report workbook closed; journal retains result");session.Batch.Confirm(File.Delete);
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.ForeignOpen=true;session.Execute(null);
        Check(File.ReadAllBytes(target).SequenceEqual(before)&&((LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook).CloseCount==0,"foreign object returned for stage is never owned or closed");session.Batch.Confirm(File.Delete);
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        iterator=session.Steps(false,()=>false).GetEnumerator();Check(iterator.MoveNext()&&iterator.MoveNext(),"stage edit fixture reaches saved stage");var editedStage=((LinkFakeApp)session.App).Workbooks.Items[1];editedStage.Saved=false;while(iterator.MoveNext()){}iterator.Dispose();
        Check(File.ReadAllBytes(target).SequenceEqual(before)&&editedStage.CloseCount==0&&WpsShared.SourcePaths.Exists(editedStage.FullName),"unexpected user edits in owned stage are preserved rather than closed or committed");editedStage.Close(false);session.Batch.Confirm(File.Delete);
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        stop=false;iterator=session.Steps(false,()=>stop).GetEnumerator();Check(iterator.MoveNext()&&iterator.MoveNext(),"stage edit cancellation reaches saved stage");editedStage=((LinkFakeApp)session.App).Workbooks.Items[1];editedStage.Saved=false;stop=true;while(iterator.MoveNext()){}iterator.Dispose();
        Check(File.ReadAllBytes(target).SequenceEqual(before)&&editedStage.CloseCount==0&&WpsShared.SourcePaths.Exists(editedStage.FullName),"cancellation preserves newly edited stage without touching original");editedStage.Close(false);session.Batch.Confirm(File.Delete);
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(target,new[]{target},a,LinkAssistantSession.BreakCommand,false);book=(LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook;
        iterator=session.Steps(false,()=>false).GetEnumerator();Check(iterator.MoveNext()&&iterator.MoveNext(),"report-write failure reaches saved break");book.Sheet.FailWrite=true;while(iterator.MoveNext()){}iterator.Dispose();
        Check(session.Batch.Checkpoints[0].Uncertain&&WpsShared.SourcePaths.Exists(session.Batch.Checkpoints[0].Backup),"partial report write failure preserves uncertain checkpoint");Reject(()=>session.Batch.Confirm(_=>{}),"report-write failure prevents clearing recovery backup");
        book.Close(false);foreach(string step in session.Steps(true,()=>false)){}Check(File.ReadAllBytes(target).SequenceEqual(before),"closed workbook restores full backup after report-write failure");session.Batch.Confirm(File.Delete);
    }
    static XmlDocument SheetXml(string file)
    {var doc=new XmlDocument();using(var stream=File.OpenRead(file))using(var zip=new ZipArchive(stream,ZipArchiveMode.Read))using(var reader=zip.GetEntry("xl/worksheets/sheet1.xml").Open())doc.Load(reader);return doc;}
    static void WriteFormulaSample(string file,string first,string second)
    {
        using(var stream=File.Open(file,FileMode.Open,FileAccess.ReadWrite))using(var zip=new ZipArchive(stream,ZipArchiveMode.Update))
        {zip.GetEntry("xl/worksheets/sheet1.xml").Delete();Part(zip,"xl/worksheets/sheet1.xml","<worksheet><c r=\"A1\"><f>'["+Path.GetFileName(first)+"]Sheet1'!A1</f><v>12345</v></c><c r=\"B1\"><f>'["+Path.GetFileName(second)+"]Sheet1'!A1</f><v>10</v></c></worksheet>");}
    }
}
public partial class LinkFakeBook
{
    public bool RefuseBreak,NoBreak,FailSaveAfterBreak;
    private byte[] pendingBreak;
    public void BreakLink(string name,int type)
    {
        if(RefuseBreak)throw new IOException("fake native BreakLink refusal");if(NoBreak){Saved=false;return;}
        string source=SourceStyle=="uri"?new Uri(name).LocalPath:name;
        using(var memory=new MemoryStream())
        {
            byte[] original=File.ReadAllBytes(FullName);memory.Write(original,0,original.Length);
            using(var zip=new ZipArchive(memory,ZipArchiveMode.Update,true))
            {
                foreach(var entry in zip.Entries.Where(e=>e.FullName.StartsWith("xl/externalLinks/_rels/")).ToList())
                {
                    var xml=new XmlDocument();using(var stream=entry.Open())xml.Load(stream);
                    if(xml.DocumentElement.ChildNodes.OfType<XmlElement>().Any(r=>string.Equals(XlsxLinkPackage.Resolve(FullName,r.GetAttribute("Target")),source,StringComparison.OrdinalIgnoreCase)))entry.Delete();
                }
                var sheet=zip.GetEntry("xl/worksheets/sheet1.xml");var doc=new XmlDocument();using(var stream=sheet.Open())doc.Load(stream);
                foreach(XmlNode formula in doc.SelectNodes("//f").Cast<XmlNode>().ToList())if(formula.InnerText.Contains("["+Path.GetFileName(source)+"]"))formula.ParentNode.RemoveChild(formula);
                sheet.Delete();using(var writer=new StreamWriter(zip.CreateEntry("xl/worksheets/sheet1.xml").Open(),new UTF8Encoding(false)))writer.Write(doc.OuterXml);
            }
            pendingBreak=memory.ToArray();
        }
        Saved=false;if(FailSaveAfterBreak){failSave=true;FailSaveAfterBreak=false;}
    }
}
