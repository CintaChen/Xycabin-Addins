using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel_Iterate;

partial class Feature099Probe
{
    const BindingFlags NamesPrivate = BindingFlags.Instance | BindingFlags.NonPublic;
    static void NamesInvoke(FileNamesForm form,string method)
    {typeof(FileNamesForm).GetMethod(method,NamesPrivate).Invoke(form,null);}
    static T NamesControl<T>(FileNamesForm form,string name) where T:Control
    {return DialogFormatting.Walk(form).OfType<T>().Single(c=>c.Name==name);}
    static ResultGrid NamesGrid(FileNamesForm form)
    {return DialogFormatting.Walk(form).OfType<ResultGrid>().Single();}
    static void NamesEntries(FileNamesForm form,List<FileNameEntry> entries,bool ocr)
    {
        typeof(FileNamesForm).GetField("entries",NamesPrivate).SetValue(form,entries);
        typeof(FileNamesForm).GetField("scannedFeatures",NamesPrivate).SetValue(form,ocr);
        NamesInvoke(form,"ShowEntries");
    }
    static void NamesPumpLayout(FileNamesForm form,int width)
    {
        form.ClientSize=new Size(width,1000);
        for(int pass=0;pass<3;pass++)
        {
            form.PerformLayout();
            foreach(var table in DialogFormatting.Walk(form).OfType<TableLayoutPanel>())table.PerformLayout();
            Application.DoEvents();
        }
    }
    static void NamesPark(FileNamesForm form,Panel parent)
    {
        parent.Size=new Size(2300,1400);form.TopLevel=false;form.FormBorderStyle=FormBorderStyle.None;
        parent.Controls.Add(form);form.Show();NamesPumpLayout(form,1200);
        Check(!IsWindowVisible(form.Handle),"file names regression remains under an invisible parent");
    }
    static int NamesCap(ResultGrid grid,DataGridViewColumn column)
    {
        var style=column.InheritedStyle;
        var font=style.Font??grid.DefaultCellStyle.Font??grid.Font;
        return TextRenderer.MeasureText(new string('0',30),font,Size.Empty,TextFormatFlags.SingleLine|TextFormatFlags.NoPadding).Width+style.Padding.Horizontal;
    }
    sealed class LayoutMockOcr:IFileOcr
    {
        internal int Calls;
        public Task<List<OcrSegment>> RecognizeAsync(string path,int page,CancellationToken token)
        {Calls++;return Task.FromResult(new List<OcrSegment>{new OcrSegment {Text="磅单号：000123",NoLocation=true,Page=page}});}
        public void Dispose(){}
    }
    static Task NamesRun(FileNamesForm form,bool retry)
    {
        var previous=SynchronizationContext.Current;
        try {SynchronizationContext.SetSynchronizationContext(null);return (Task)typeof(FileNamesForm).GetMethod("Run",NamesPrivate).Invoke(form,new object[]{retry});}
        finally {SynchronizationContext.SetSynchronizationContext(previous);}
    }
    static void NamesWorkflowChecks()
    {
        string folder=Path.Combine(sandbox,"layout-scan");Directory.CreateDirectory(folder);
        using(var image=new Bitmap(100,60))image.Save(Path.Combine(folder,"one.png"));
        var book=new FakeWorkbook();var sheet=new FakeSheet(book);book.Worksheets.Items.Add(sheet);
        using(var form=new FileNamesForm(new FileNameWorksheet(new FakeApplication(sheet))))using(var parent=new Panel())
        {
            NamesPark(form,parent);var grid=NamesGrid(form);var feature=grid.Columns["feature"];
            Check(grid.Columns.Count==6&&grid.Columns.Cast<DataGridViewColumn>().All(c=>c.Visible),"all six file result headers exist before a scan");
            Check(grid.Columns.Cast<DataGridViewColumn>().Select(c=>c.HeaderText).SequenceEqual(new[]{"文件提取名称","文件后缀","文件路径","文件OCR识别","候选","识别状态"}),"initial file result headers retain the six semantic names");
            NamesControl<TextBox>(form,"folderInput").Text=folder;Pump(NamesRun(form,false));
            var row=grid.Rows[0];var entry=(FileNameEntry)row.Tag;
            Check(grid.Columns.Cast<DataGridViewColumn>().All(c=>c.Visible)&&row.Cells["feature"].ReadOnly,"non-OCR scan keeps the six headers and a read-only result");
            Check(grid.Columns.Cast<DataGridViewColumn>().All(c=>c.ReadOnly),"non-OCR completion keeps source identity and status columns read-only");
            Check(Convert.ToString(row.Cells["feature"].Value)==""&&Convert.ToString(row.Cells["status"].Value)=="未启用OCR","non-OCR scan shows blank recognition and an explicit disabled status");
            Check(Convert.ToString(row.Cells["candidates"].Value)=="—"&&entry.Fields.Count==0,"non-OCR rows expose no field candidates");
            var plain=FileNameEngine.Values(new[]{entry},false);
            Check(plain.GetLength(1)==3&&Convert.ToString(plain[0,2])=="文件路径","non-OCR worksheet output remains three columns");
            using(var destination=new FileNameWorksheet(new FakeApplication(sheet)))
            {Check(destination.Write(new[]{entry},1,false,()=>true)&&sheet.Cells[1,4].Value2==null,"non-OCR actual writer does not add an OCR worksheet column");}
            form.SetPrefixRules(new List<OcrPrefixRule>{new OcrPrefixRule {Text="磅单号："}});
            NamesControl<CheckBox>(form,"useOcr").Checked=true;
            NamesControl<TextBox>(form,"apiKey").Text="mock";NamesControl<TextBox>(form,"secretKey").Text="mock";
            var client=new LayoutMockOcr();form.OcrFactory=(key,secret,mode)=>client;form.ConfirmUpload=_=>true;
            Pump(NamesRun(form,false));entry=(FileNameEntry)grid.Rows[0].Tag;
            Check(client.Calls==1&&entry.Feature=="000123"&&!grid.Rows[0].Cells["feature"].ReadOnly,"OCR scan resolves the editable fourth result without extra uploads");
            Check(grid.Columns.Cast<DataGridViewColumn>().Where(c=>c.Name!="feature").All(c=>c.ReadOnly),"OCR completion restores read-only source identity and status columns");
            Check(Convert.ToString(grid.Rows[0].Cells["candidates"].Value)=="核对 / 留空","OCR rows retain the complete review button text");
            var identified=FileNameEngine.Values(new[]{entry},true);
            Check(identified.GetLength(1)==4&&Convert.ToString(identified[1,3])=="000123","OCR worksheet output remains four columns and retains leading zeros");
            using(var destination=new FileNameWorksheet(new FakeApplication(sheet)))
            {Check(destination.Write(new[]{entry},1,true,()=>true)&&Convert.ToString(sheet.Cells[2,4].Value2)=="000123","OCR actual writer preserves the fourth value as text");}
            int manual=NamesCap(grid,feature)+140;feature.Width=manual;
            Pump(NamesRun(form,true));Check(feature.Width==manual&&client.Calls==1,"retry of already resolved OCR rows preserves the manually widened column");
            entry.ManualFeature=true;entry.Feature="manual whole row";
            form.SetPrefixRules(new List<OcrPrefixRule>{new OcrPrefixRule {Text="磅单号："}});
            Check(feature.Width==manual&&entry.Feature=="manual whole row","rule recomputation preserves both manual whole-row text and its column width");
            Pump(NamesRun(form,false));Check(feature.Width!=manual&&client.Calls==2,"a new scan begins a fresh automatic column-fit cycle");
            form.Close();
        }
    }
    static void NamesWidthChecks()
    {
        using(var form=new FileNamesForm(null))using(var parent=new Panel())
        {
            NamesPark(form,parent);var grid=NamesGrid(form);var feature=grid.Columns["feature"];
            Check(grid.AutoSizeColumnsMode==DataGridViewAutoSizeColumnsMode.None&&grid.AllowUserToResizeColumns,"result grid uses manual column sizes with user resizing enabled");
            Check(grid.Columns.Cast<DataGridViewColumn>().All(c=>c.AutoSizeMode!=DataGridViewAutoSizeColumnMode.Fill),"file result columns do not inherit continuous fill sizing");
            Check(grid.ScrollBars==ScrollBars.Both,"file results permit horizontal and vertical scrolling");
            var item=new FileNameEntry {Name="sample",Extension="png",Path="C:\\fixture\\sample.png",Feature="short",Status="已提取",Resolved=true};
            NamesInvoke(form,"BeginColumnFitCycle");NamesEntries(form,new List<FileNameEntry>{item},true);NamesInvoke(form,"FitResultColumns");
            int shortWidth=feature.Width;int cap=NamesCap(grid,feature);
            Check(feature.MinimumWidth>=Ui.Px(grid,50)&&shortWidth>=feature.MinimumWidth&&shortWidth<=Math.Max(feature.MinimumWidth,cap)+8,"automatic short result respects the complete header minimum and bounded content width");
            item.Feature=new string('0',40);NamesEntries(form,new List<FileNameEntry>{item},true);NamesInvoke(form,"FitResultColumns");
            Check(feature.Width>shortWidth&&feature.Width<=cap+8&&feature.Width>=cap-8,"automatic long result reaches the thirty-zero content cap");
            Check(Convert.ToString(grid.Rows[0].Cells["feature"].Value)==item.Feature,"bounded columns retain the entire underlying long value");
            Check(feature.HeaderCell.InheritedStyle.WrapMode==DataGridViewTriState.False,"long result headers remain single-line with a complete minimum");
            int buttonNeed=TextRenderer.MeasureText("核对 / 留空",grid.Columns["candidates"].InheritedStyle.Font??grid.Font).Width+grid.Columns["candidates"].InheritedStyle.Padding.Horizontal;
            Check(grid.Columns["candidates"].Width>=buttonNeed,"candidate action column fits the complete action caption");
            feature.Width=cap+160;int widened=feature.Width;
            NamesInvoke(form,"FitResultColumns");Ui.AutoFitColumns(grid);NamesInvoke(form,"ShowEntries");NamesPumpLayout(form,1800);
            Check(feature.Width==widened,"manual width above the automatic cap survives fit, shared styling, refresh and resize");
            feature.Width=feature.MinimumWidth;int narrow=feature.Width;NamesInvoke(form,"FitResultColumns");Ui.AutoFitColumns(grid);
            Check(feature.Width==narrow,"manual narrowing reaches the minimum and is not snapped back by fitting");
            NamesInvoke(form,"BeginColumnFitCycle");NamesInvoke(form,"FitResultColumns");
            Check(feature.Width>narrow&&feature.Width<=cap+8,"new fit cycle releases the prior manual override");
            var target=DialogFormatting.Register(form,"file.names").Single(t=>t.Control==grid&&!t.Caption);
            var custom=new CustomizationData();custom.Formats["file.names::"+target.Id]=new FormatRule {Size=14,Family=grid.Font.Name};
            feature.Width=cap+160;int customizedWidth=feature.Width;DialogFormatting.Apply(form,"file.names",custom);
            Check(feature.Width==customizedWidth,"format editor font application cannot overwrite a manually resized result column");
        }
        using(var merge=new MergeHeadersForm())
        {
            Ui.AutoFitColumns(merge.Map);
            Check(merge.Map.Columns[merge.Map.Columns.Count-1].AutoSizeMode==DataGridViewAutoSizeColumnMode.Fill,"ordinary merge result grid retains its existing final fill column");
        }
        foreach(float factor in new[]{1f,1.25f,1.5f,2f})
        using(var form=new FileNamesForm(null))using(var parent=new Panel())using(var font=Ui.MakeFont(10f*factor,FontStyle.Regular,form.Font))
        {
            NamesPark(form,parent);var grid=NamesGrid(form);grid.DefaultCellStyle.Font=font;
            NamesInvoke(form,"BeginColumnFitCycle");NamesEntries(form,new List<FileNameEntry>{new FileNameEntry {Name="sample",Feature=new string('0',80),Status="已提取"}},true);NamesInvoke(form,"FitResultColumns");
            int cap=NamesCap(grid,grid.Columns["feature"]);
            Check(grid.Columns["feature"].Width<=cap+8&&grid.Columns["feature"].Width>=cap-8,"thirty-zero cap uses current font at synthetic font scale "+factor);
            Check(grid.Columns["feature"].MinimumWidth>=Ui.Px(grid,50),"complete header minimum includes the actual display DPI floor "+factor);
            Console.WriteLine("File results synthetic font scale "+factor+"; display DPI is unchanged");
        }
    }
    static void NamesLayoutChecks()
    {
        using(var form=new FileNamesForm(null))using(var parent=new Panel())
        {
            NamesPark(form,parent);var options=NamesControl<TableLayoutPanel>(form,"fileOptions");var help=NamesControl<Label>(form,"targetSummary");var grid=NamesGrid(form);
            NamesPumpLayout(form,900);int narrowHeight=help.Height,narrowOptions=options.Height,narrowTop=grid.Top;
            foreach(int width in new[]{1200,1800,900,1800,1200})
            {
                NamesPumpLayout(form,width);
                Check(options.ClientRectangle.Bottom-help.Bottom<=help.Margin.Bottom+options.Padding.Bottom+4,"explanation row has no stale lower whitespace at width "+width);
                Check(grid.Top-options.Bottom<=grid.Margin.Top+options.Margin.Bottom+4,"grid follows the collapsed configuration container at width "+width);
                Check(options.GetRowHeights().All(h=>h>0)&&grid.Height>250,"all semantic configuration rows and result area remain usable at width "+width);
            }
            NamesPumpLayout(form,1800);
            Check(help.Height<narrowHeight&&options.Height<narrowOptions&&grid.Top<narrowTop,"wide window shortens wrapped explanation and moves the result grid upward");
            Check(help.Height<=TextRenderer.MeasureText("Ag",help.Font).Height+4,"widest default explanation contracts to one text line");
            NamesPumpLayout(form,900);Check(help.Height>=narrowHeight&&grid.Top>=narrowTop,"narrowing restores the required wrapped explanation height");
            NamesControl<ComboBox>(form,"ocrProvider").SelectedIndex=1;NamesPumpLayout(form,1800);
            Check(help.Text.Contains("阿里云不支持 PDF")&&options.ClientRectangle.Bottom-help.Bottom<=help.Margin.Bottom+4,"provider explanation changes also reclaim obsolete row height");
        }
        foreach(float size in new[]{12f,18f})
        using(var form=new FileNamesForm(null))using(var parent=new Panel())
        {
            NamesPark(form,parent);var targets=DialogFormatting.Register(form,"file.names");
            var api=NamesControl<TextBox>(form,"apiKey");var help=NamesControl<Label>(form,"targetSummary");
            var data=new CustomizationData();
            foreach(var target in targets.Where(t=>t.Control is Label||t.Control is Button||t.Control is ComboBox||t.Control is TextBox||t.Control is CheckBox))
                data.Formats["file.names::"+target.Id]=new FormatRule {Size=size,Family=target.OriginalFont.Name,Style=target.OriginalFont.Bold?1:0};
            string key="file.names::"+targets.Single(t=>t.Control==api).Id;data.Formats[key].Width96=340;
            var restored=LocalCustomization.Decode(LocalCustomization.Encode(data));DialogFormatting.Apply(form,"file.names",restored);NamesPumpLayout(form,900);
            var options=NamesControl<TableLayoutPanel>(form,"fileOptions");int oldHeight=options.Height;
            Check(api.Width==Ui.Px(api,340)&&DialogControlSizing.HasWidth(api),"saved input width survives customized font/layout roundtrip at "+size);
            foreach(int width in new[]{1800,900,1800})
            {
                NamesPumpLayout(form,width);int previousBottom=0;
                foreach(Control row in options.Controls.Cast<Control>().OrderBy(c=>options.GetRow(c)))
                {Check(row.Top>=previousBottom,"configured rows do not overlap after width/font change "+width+" / "+size);previousBottom=row.Bottom;}
                Check(options.ClientRectangle.Bottom-help.Bottom<=help.Margin.Bottom+4,"larger font explanation releases unused row height at "+width+" / "+size);
                Check(api.Width==Ui.Px(api,340),"reflow preserves the explicit input width at "+width+" / "+size);
            }
            Check(options.Height<oldHeight,"larger fonts still allow the configuration region to contract at "+size);
        }
    }
    static void FileNamesLayoutChecks()
    {
        int before=checks;NamesWorkflowChecks();NamesWidthChecks();NamesLayoutChecks();
        Console.WriteLine("File names column sizing and collapsing layout: "+(checks-before)+" PASS; hidden fake worksheet/mock OCR, synthetic font scaling, no WPS or cloud");
    }
}
