using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel_Iterate;

partial class Feature099Probe
{
    sealed class CombinedMockOcr : IFileOcr
    {
        internal int Calls;
        public Task<List<OcrSegment>> RecognizeAsync(string path,int page,CancellationToken token)
        {Calls++;return Task.FromResult(new List<OcrSegment>{new OcrSegment {Text="磅单号：00123 毛重：10kg 销售订单号：00987",NoLocation=true,Page=page}});}
        public void Dispose(){}
    }
    static void MultiplePrefixWorkflow(List<OcrPrefixRule> rules)
    {
        string folder=Path.Combine(sandbox,"multiple-prefix-upload");Directory.CreateDirectory(folder);using(var image=new Bitmap(100,60))image.Save(Path.Combine(folder,"one.png"));
        using(var form=new FileNamesForm(new FileNameWorksheet(new FakeApplication(new FakeSheet(new FakeWorkbook())))))using(var parent=new Panel())
        {
            form.TopLevel=false;parent.Controls.Add(form);form.Show();var all=DialogFormatting.Walk(form).ToList();
            form.SetPrefixRules(rules);all.OfType<TextBox>().Single(c=>c.Name=="folderInput").Text=folder;all.OfType<TextBox>().Single(c=>c.Name=="apiKey").Text="mock";all.OfType<TextBox>().Single(c=>c.Name=="secretKey").Text="mock";all.OfType<CheckBox>().Single(c=>c.Name=="useOcr").Checked=true;
            var client=new CombinedMockOcr();form.OcrFactory=(key,secret,mode)=>client;form.ConfirmUpload=_=>true;
            var context=SynchronizationContext.Current;Task run;try{SynchronizationContext.SetSynchronizationContext(null);run=(Task)typeof(FileNamesForm).GetMethod("Run",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,new object[]{false});}finally{SynchronizationContext.SetSynchronizationContext(context);}
            Pump(run);var grid=all.OfType<ResultGrid>().Single();var entry=(FileNameEntry)grid.Rows[0].Tag;
            Check(client.Calls==1&&entry.Feature=="00123-10kg-00987","three prefixes use one actual OCR workflow request");
            form.SetPrefixRules(rules.AsEnumerable().Reverse().ToList());Check(client.Calls==1&&entry.Feature=="00987-10kg-00123","changing rules recomposes cached workflow without upload");
            var values=FileNameEngine.Values(new[]{entry},true);Check((string)values[1,3]=="00987-10kg-00123","fourth OCR result column receives composed text");
            var kinds=rules.Select(r=>r.Copy()).ToList();kinds[1].Kinds=OcrCharacterKinds.Digits;form.SetPrefixRules(kinds);Check(client.Calls==1&&entry.Feature=="00123-10-00987","per-prefix character change uses cached workflow without upload");
            form.SetPrefixRules(rules);Check(client.Calls==1&&entry.Feature=="00123-10kg-00987","restoring all character kinds recomposes from raw OCR cache");
            entry.ManualFeature=true;entry.Feature="manual row";var disabled=rules.Select(r=>r.Copy()).ToList();foreach(var rule in disabled)rule.Enabled=false;
            form.SetPrefixRules(disabled);Check(entry.Resolved&&entry.Feature=="manual row","disabling all prefixes preserves whole-row manual text");form.SetPrefixRules(rules);Check(entry.Feature=="manual row"&&client.Calls==1,"reenabling prefixes preserves whole-row override without uploading");
            entry.ExplicitBlank=true;entry.Feature="";form.SetPrefixRules(disabled);Check(entry.Resolved&&entry.Status=="已明确留空","disabling all prefixes preserves explicit blank");form.Close();
        }
    }
    static void MultiplePrefixes()
    {
        CharacterKinds();
        var rules=OcrCombined.ExampleRules();
        var text=new List<OcrSegment>{new OcrSegment {Text="磅单号：0012345、毛重：10kg；销售订单号：0098765",NoLocation=true,Page=1}};
        var entry=new FileNameEntry {OcrSegments=text};OcrCombined.Apply(entry,rules);
        Check(entry.Resolved&&entry.Feature=="0012345-10kg-0098765","multiple fields slice labels, preserve zeros and units");
        OcrCombined.Apply(entry,rules.AsEnumerable().Reverse().ToList());Check(entry.Feature=="0098765-10kg-0012345","prefix order controls composition");
        var all=rules.Select(r=>r.Copy()).ToList();all[1].Enabled=false;OcrCombined.Apply(entry,all);Check(entry.Feature=="0012345-0098765","disabled label remains a field boundary");
        var partial=new FileNameEntry {OcrSegments=new List<OcrSegment>{new OcrSegment {Text="磅单号：0012345",NoLocation=true,Page=1},new OcrSegment {Text="销售订单号：0098765",NoLocation=true,Page=2}}};OcrCombined.Apply(partial,rules);
        Check(!partial.Resolved&&partial.Feature==""&&partial.Status.Contains("毛重"),"missing field blocks partial composition");
        partial.Fields[1].Chosen="10kg";partial.Fields[1].Manual=true;OcrCombined.Compose(partial);Check(partial.Feature=="0012345-10kg-0098765","manual supplement completes composition");
        OcrCombined.Apply(partial,rules.AsEnumerable().Reverse().ToList());Check(partial.Feature=="0098765-10kg-0012345","cached reevaluation retains manual value");
        text.Add(new OcrSegment {Text="毛重：11kg",NoLocation=true,Page=2});OcrCombined.Apply(entry,rules);Check(!entry.Resolved&&entry.Fields[1].Candidates.Count==2,"multiple candidates require confirmation");
        entry.Fields[1].Chosen="11kg";entry.Fields[1].Manual=true;OcrCombined.Compose(entry);Check(entry.Feature=="0012345-11kg-0098765","chosen candidate controls composition");
        OcrCombined.Apply(entry,rules);Check(entry.Feature=="0012345-11kg-0098765","explicit choice retained without another upload");
        entry.ExplicitBlank=true;OcrCombined.Apply(entry,rules);Check(entry.Resolved&&entry.Feature==""&&entry.Status=="已明确留空","explicit blank remains resolved");
        entry.ExplicitBlank=false;entry.ManualFeature=true;entry.Feature="手工整行";OcrCombined.Apply(entry,rules);Check(entry.Feature=="手工整行","manual complete override remains intact");
        var overlapping=new List<OcrPrefixRule>{new OcrPrefixRule {Text="订单号："},new OcrPrefixRule {Text="销售订单号："}};
        var matched=OcrCombined.Extract(new[]{new OcrSegment {Text="销售订单号：0098",NoLocation=true,Page=1}},overlapping);Check(matched[0].Candidates.Count==0&&matched[1].Chosen=="0098","longer overlapping label wins");
        Check(OcrCombined.Extract(new[]{new OcrSegment {Text="磅单号：;001,",NoLocation=true,Page=1}},new List<OcrPrefixRule>{rules[0]})[0].Chosen==";001,","single-prefix punctuation remains compatible");
        var located=new[]{new OcrSegment {Text="磅单号：",Left=0,Width=60,Height=20,Page=1},new OcrSegment {Text="001",Left=65,Width=30,Height=20,Page=1},new OcrSegment {Text="毛重：10kg",Left=140,Width=100,Height=20,Page=1},new OcrSegment {Text="销售订单号：009",Left=0,Top=50,Width=200,Height=20,Page=1}};
        var locatedEntry=new FileNameEntry {OcrSegments=located.ToList()};OcrCombined.Apply(locatedEntry,rules);Check(locatedEntry.Feature=="001-10kg-009","positioned fragments retain row boundaries");
        var isolated=OcrCombined.Extract(new[]{new OcrSegment {Text="磅单号：",Height=20,Page=1},new OcrSegment {Text="001",Height=20,Page=2}},rules);Check(isolated[0].Candidates.Count==0,"pages never concatenated");
        var repeated=rules.Select(r=>r.Copy()).ToList();repeated.Add(new OcrPrefixRule {Text="磅 单号"});Reject(()=>OcrPrefixStore.Validate(repeated,true),"duplicate optional-colon prefixes rejected");Reject(()=>OcrPrefixStore.Validate(new List<OcrPrefixRule>(),true),"OCR requires enabled prefixes");
        var initial=OcrPrefixStore.Read();var saved=OcrPrefixStore.Save(initial,rules);Check(saved.Writable&&saved.Rules.Select(r=>r.Text).SequenceEqual(rules.Select(r=>r.Text)),"rules persist atomically in isolated XML");Reject(()=>OcrPrefixStore.Save(initial,rules),"stale editor cannot overwrite newer rules");
        var flipped=saved.Rules.AsEnumerable().Reverse().Select(r=>r.Copy()).ToList();flipped[0].Enabled=false;var savedAgain=OcrPrefixStore.Save(saved,flipped);Check(!savedAgain.Rules[0].Enabled&&savedAgain.Rules[1].Id==rules[1].Id,"order enabled state and identities survive reopening");
        using(var form=(OcrPrefixesForm)DialogPreviewFactory.Create("file.ocrprefixes",DialogCatalog.Find("file.ocrprefixes").Default))
        {string before=WpsShared.SourcePaths.ReadAllText(OcrPrefixStore.PathForUser);form.Submit();Check(form.Result==null&&WpsShared.SourcePaths.ReadAllText(OcrPrefixStore.PathForUser)==before&&!form.Visible,"prefix preview does not save");}
        using(var form=(OcrFieldsForm)DialogPreviewFactory.Create("file.ocrfields",DialogCatalog.Find("file.ocrfields").Default))
        {form.Submit();Check(form.Result==null&&!form.Visible&&!((Control)form.AcceptButton).Enabled,"field preview cannot accept business action");}
        int saves=0;using(var form=new OcrPrefixesForm(rules,rows=>{saves++;return "";}))
        {
            Check(new[]{"chinese","digits","letters","symbols"}.All(name=>form.Grid.Columns[name] is DataGridViewCheckBoxColumn),"prefix editor exposes four independent character checkbox columns");
            Check(form.Grid.Rows.Cast<DataGridViewRow>().All(row=>Enumerable.Range(2,4).All(i=>Convert.ToBoolean(row.Cells[i].Value))),"existing prefix rows default all character kinds checked");
            form.Grid.Rows[1].Cells["chinese"].Value=false;form.Grid.Rows[1].Cells["letters"].Value=false;form.Grid.Rows[1].Cells["symbols"].Value=false;
            form.Grid.CurrentCell=form.Grid.Rows[1].Cells[1];typeof(OcrPrefixesForm).GetMethod("MoveRule",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,new object[]{-1});
            Check(((OcrPrefixRule)form.Grid.Rows[0].Tag).Id==rules[1].Id,"prefix editor reorders model identities with rows");form.Submit();Check(form.Result[0].Id==rules[1].Id&&saves==1,"prefix editor submits ordered model once");Check(form.Result[0].Kinds==OcrCharacterKinds.Digits&&form.Result[1].Kinds==OcrCharacterKinds.All,"prefix editor submits character selections per reordered identity");
        }
        using(var form=new OcrPrefixesForm(rules,rows=>{saves++;return "";}))
        {object[] args={new Message(),Keys.Escape};typeof(OcrPrefixesForm).GetMethod("ProcessCmdKey",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,args);Check(form.DialogResult==DialogResult.Cancel&&form.Result==null&&saves==1,"prefix Esc cancels without saving");}
        var missingFields=OcrCombined.Extract(new[]{new OcrSegment {Text="磅单号：001 销售订单号：009",NoLocation=true,Page=1}},rules);
        using(var form=new OcrFieldsForm(missingFields))
        {form.Submit();Check(form.Result==null&&!form.IsDisposed,"field editor rejects incomplete values without accepting");form.Grid.Rows[1].Cells[2].Value="10kg";form.Submit();var accepted=new FileNameEntry {Fields=form.Result};OcrCombined.Compose(accepted);Check(accepted.Feature=="001-10kg-009"&&accepted.Fields[1].Manual,"field editor accepts explicit missing-value supplement");}
        using(var form=new OcrFieldsForm(missingFields))
        {object[] args={new Message(),Keys.Escape};typeof(OcrFieldsForm).GetMethod("ProcessCmdKey",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,args);Check(form.DialogResult==DialogResult.Cancel&&form.Result==null,"field Esc cancels without composing");}
        using(var form=new OcrPrefixesForm(rules,rows=>{saves++;return "";}))
        {foreach(string name in new[]{"chinese","digits","letters","symbols"})form.Grid.Rows[0].Cells[name].Value=false;form.Submit();Check(form.Result==null&&saves==1&&!form.IsDisposed,"prefix editor rejects an enabled row with no character kinds");form.Grid.Rows[0].Cells["enabled"].Value=false;form.Submit();Check(form.Result!=null&&form.Result[0].Kinds==OcrCharacterKinds.None&&saves==2,"disabled prefix may remember no selected character kinds");}
        CharacterRuleStorage(rules);
        string path=OcrPrefixStore.PathForUser;File.WriteAllText(path,"<OcrPrefixRules version=\"3\"/>");var future=OcrPrefixStore.Read();Reject(()=>OcrPrefixStore.Save(future,rules),"future settings preserved");
        File.WriteAllText(path,"<!DOCTYPE x [<!ENTITY y SYSTEM 'file:///not-read'>]><OcrPrefixRules version=\"1\">&y;</OcrPrefixRules>");Check(!OcrPrefixStore.Read().Writable,"DTD settings rejected");File.Delete(path);
        MultiplePrefixWorkflow(rules);
    }
    static void CharacterKinds()
    {
        for(int mask=0;mask<=15;mask++)
        {
            var kinds=(OcrCharacterKinds)mask;
            string expected=((kinds&OcrCharacterKinds.Letters)!=0?"X":"")+((kinds&OcrCharacterKinds.Digits)!=0?"187":"")+((kinds&OcrCharacterKinds.Chinese)!=0?"向宇":"")+((kinds&OcrCharacterKinds.Symbols)!=0?"##":"");
            Check(OcrCombined.FilterCharacters("X187向宇##",kinds)==expected,"character mask "+mask+" retains only selected categories in order");
        }
        string extraHan=char.ConvertFromUtf32(0x20000),emoji=char.ConvertFromUtf32(0x1f600),full="\tＸｘ１８７ 向宇"+extraHan+"##"+emoji+"\r\n";
        Check(OcrCombined.FilterCharacters(full,OcrCharacterKinds.Chinese)=="向宇"+extraHan,"supplementary Han uses the existing file-name Chinese range without splitting surrogates");
        Check(OcrCombined.FilterCharacters(full,OcrCharacterKinds.Digits)=="１８７","fullwidth digits are numeric and keep their original characters");
        Check(OcrCombined.FilterCharacters(full,OcrCharacterKinds.Letters)=="Ｘｘ","fullwidth uppercase and lowercase English letters are letters");
        Check(OcrCombined.FilterCharacters(full,OcrCharacterKinds.Symbols)=="##"+emoji,"emoji is one complete symbol and whitespace is excluded");
        Check(OcrCombined.FilterCharacters(null,OcrCharacterKinds.All)==""&&OcrCombined.FilterCharacters(" \t\r\n　",OcrCharacterKinds.All)=="","empty and Unicode whitespace produce no candidate characters");
        var one=new List<OcrPrefixRule>{new OcrPrefixRule {Text="单号：",Kinds=OcrCharacterKinds.Digits}};
        var entry=new FileNameEntry {OcrSegments=new List<OcrSegment>{new OcrSegment {Text="单号：X187向宇##",NoLocation=true,Page=1}}};OcrCombined.Apply(entry,one);
        Check(entry.Feature=="187"&&entry.Fields[0].Prefix=="单号：","filter applies to extracted value and never removes the field marker");
        one[0].Kinds=OcrCharacterKinds.Symbols;OcrCombined.Apply(entry,one);Check(entry.Feature=="##","only symbols produces the requested sample result");
        one[0].Kinds=OcrCharacterKinds.Digits|OcrCharacterKinds.Letters;OcrCombined.Apply(entry,one);Check(entry.Feature=="X187","digits plus letters produces the requested sample result");
        one[0].Kinds=OcrCharacterKinds.Digits;entry.OcrSegments.Add(new OcrSegment {Text="单号：Y187###",NoLocation=true,Page=2});OcrCombined.Apply(entry,one);Check(entry.Resolved&&entry.Feature=="187"&&entry.Fields[0].Candidates.Count==1,"filtering precedes candidate deduplication and automatic unique selection");
        entry.OcrSegments.Add(new OcrSegment {Text="单号：188",NoLocation=true,Page=3});OcrCombined.Apply(entry,one);Check(!entry.Resolved&&entry.Fields[0].Candidates.SequenceEqual(new[]{"187","188"}),"different filtered values still require candidate review");
        var rules=new List<OcrPrefixRule>{new OcrPrefixRule {Text="单号：",Kinds=OcrCharacterKinds.Digits},new OcrPrefixRule {Text="姓名：",Kinds=OcrCharacterKinds.Chinese}};
        var mixed=new FileNameEntry {OcrSegments=new List<OcrSegment>{new OcrSegment {Text="单号：X001##、姓名：abc向宇98",NoLocation=true,Page=1}}};OcrCombined.Apply(mixed,rules);
        Check(mixed.Feature=="001-向宇","different per-prefix filters preserve leading zeroes and the fixed combination hyphen");
        mixed.Fields[0].Manual=true;mixed.Fields[0].Chosen="人工X001##";rules[0].Kinds=OcrCharacterKinds.Symbols;OcrCombined.Apply(mixed,rules);Check(mixed.Feature=="人工X001##-向宇","rule changes preserve manually chosen field values without filtering");
        var empty=new FileNameEntry {OcrSegments=new List<OcrSegment>{new OcrSegment {Text="单号：向宇",NoLocation=true,Page=1}}};OcrCombined.Apply(empty,one);
        Check(!empty.Resolved&&empty.Feature==""&&empty.Fields[0].FilteredEmpty&&empty.Status.Contains("筛选后无可用字符"),"all filtered characters removed leaves a reviewable empty field with a precise reason");
        Check(empty.Fields[0].Copy().FilteredEmpty,"copy retains the empty-after-filter diagnostic");
        using(var form=new OcrFieldsForm(new[]{new OcrFieldResult {RuleId=one[0].Id,Prefix=one[0].Text,Chosen="X187",Candidates=new List<string>{"X187"}}}))
        {
            typeof(DataGridView).GetMethod("OnCellEndEdit",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form.Grid,new object[]{new DataGridViewCellEventArgs(1,0)});form.Submit();
            var selected=new FileNameEntry {Fields=form.Result,OcrSegments=new List<OcrSegment>{new OcrSegment {Text="单号：X187",NoLocation=true,Page=1}}};OcrCombined.Apply(selected,one);
            Check(selected.Feature=="X187"&&selected.Fields[0].Manual,"explicitly selecting the same unique candidate remains a manual choice after filtering changes");
        }
        using(var form=new OcrFieldsForm(empty.Fields))
        {form.Grid.Rows[0].Cells[2].Value="人工X001##";form.Submit();empty.Fields=form.Result;OcrCombined.Compose(empty);Check(empty.Resolved&&empty.Feature=="人工X001##"&&empty.Fields[0].Manual,"manual supplement of a filtered-empty field is not filtered again");OcrCombined.Apply(empty,one);Check(empty.Feature=="人工X001##","manual supplement survives cached reevaluation under the same label");}
        empty.ExplicitBlank=true;empty.Feature="";OcrCombined.Apply(empty,one);Check(empty.Resolved&&empty.Status=="已明确留空","explicit blank overrides the filtered-empty diagnostic");
        empty.ExplicitBlank=false;empty.ManualFeature=true;empty.Feature="manual whole row##";one[0].Kinds=OcrCharacterKinds.Chinese;OcrCombined.Apply(empty,one);Check(empty.Feature=="manual whole row##","whole-row manual override is not filtered when rules change");
        var missing=OcrCombined.Extract(new[]{new OcrSegment {Text="没有这个前缀",NoLocation=true,Page=1}},one)[0];Check(!missing.FilteredEmpty,"missing marker is distinct from an existing value removed by filtering");
        var invalid=one[0].Copy();invalid.Kinds=OcrCharacterKinds.None;Reject(()=>OcrPrefixStore.Validate(new[]{invalid},false),"enabled prefix requires at least one character type");invalid.Enabled=false;OcrPrefixStore.Validate(new[]{invalid},false);Check(invalid.Copy().Kinds==OcrCharacterKinds.None,"disabled prefix and copies retain zero selected categories");invalid.Kinds=(OcrCharacterKinds)16;Reject(()=>OcrPrefixStore.Validate(new[]{invalid},false),"unknown character flags are rejected even for disabled rules");
    }
    static void CharacterRuleStorage(List<OcrPrefixRule> rules)
    {
        string path=OcrPrefixStore.PathForUser;
        string legacy="<OcrPrefixRules version=\"1\"><Prefix id=\""+rules[0].Id+"\" enabled=\"true\" text=\"磅单号：\"/><Prefix id=\""+rules[1].Id+"\" enabled=\"false\" text=\"毛重：\"/></OcrPrefixRules>";
        File.WriteAllText(path,legacy);var migrated=OcrPrefixStore.Read();Check(migrated.Writable&&migrated.Rules.Count==2&&migrated.Rules.All(r=>r.Kinds==OcrCharacterKinds.All),"version 1 settings migrate in memory to all character categories");Check(WpsShared.SourcePaths.ReadAllText(path)==legacy,"reading old rules does not rewrite the user's file");
        migrated.Rules[0].Kinds=OcrCharacterKinds.Digits|OcrCharacterKinds.Letters;migrated.Rules[1].Kinds=OcrCharacterKinds.None;
        var updated=OcrPrefixStore.Save(migrated,migrated.Rules);Check(updated.Writable&&updated.Rules[0].Kinds==(OcrCharacterKinds.Digits|OcrCharacterKinds.Letters)&&updated.Rules[1].Kinds==OcrCharacterKinds.None,"version 2 settings preserve per-prefix categories and disabled zero mask after reopening");Check(WpsShared.SourcePaths.ReadAllText(path).Contains("version=\"2\"")&&WpsShared.SourcePaths.ReadAllText(path).Contains("kinds=\"6\""),"saving migrated settings writes version 2 with explicit category masks");
        foreach(string attr in new[]{""," kinds=\"abc\""," kinds=\"16\""," kinds=\"-1\""," kinds=\"0\""})
        {
            string bad="<OcrPrefixRules version=\"2\"><Prefix id=\""+rules[0].Id+"\" enabled=\"true\" text=\"磅单号：\""+attr+"/></OcrPrefixRules>";File.WriteAllText(path,bad);var rejected=OcrPrefixStore.Read();Check(!rejected.Writable&&rejected.Rules.Count==0,"invalid version 2 character setting "+attr+" is not accepted");Reject(()=>OcrPrefixStore.Save(rejected,rules),"invalid character setting original cannot be overwritten");Check(WpsShared.SourcePaths.ReadAllText(path)==bad,"invalid character setting bytes are preserved");
        }
    }
}
