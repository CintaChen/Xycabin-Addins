using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
#if WPS_WORD
using WordUpdater;
#else
using Excel_Iterate;
#endif

partial class Feature099Probe
{
    static void AssertGridHeaders(DataGridView grid,string label)
    {
        Check(!grid.EnableHeadersVisualStyles,label+" uses owned header styles");
        foreach(DataGridViewColumn column in grid.Columns)
        {
            var style=column.HeaderCell.InheritedStyle;
            Check(style.WrapMode==DataGridViewTriState.False,label+" / "+column.HeaderText+" stays on one line");
            Check(style.Alignment==DataGridViewContentAlignment.MiddleCenter,label+" / "+column.HeaderText+" is centered");
            // Compare with WinForms' own header-cell measurement, not the Ui helper.
            int preferred=column.GetPreferredWidth(DataGridViewAutoSizeColumnMode.ColumnHeader,true);
            Check(column.MinimumWidth>=preferred,label+" / "+column.HeaderText+" minimum includes the complete native header");
            Check(column.Width>=column.MinimumWidth,label+" / "+column.HeaderText+" fill/manual width respects the header floor");
            int height=TextRenderer.MeasureText("Ag中文",style.Font??grid.Font).Height+style.Padding.Vertical;
            Check(grid.ColumnHeadersHeight>=height,label+" / "+column.HeaderText+" has enough single-line height");
            Check(grid.ColumnHeadersHeight>=column.HeaderCell.PreferredSize.Height,label+" / "+column.HeaderText+" height includes the native header-cell measurement");
        }
    }
    static void RenderGridHeaders(Form form,string id,string suffix)
    {
        string destination=Environment.GetEnvironmentVariable("XY_GRID_PROBE_RENDER_DIR");
        if(string.IsNullOrWhiteSpace(destination))return;
        Directory.CreateDirectory(destination);
        using(var bitmap=new Bitmap(form.Width,form.Height))
        {form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(destination,id+"-"+suffix+".png"));}
    }
    static void GridHeaderFontScaleChecks(Form form,string id,DataGridView[] grids)
    {
        var targets=DialogFormatting.Register(form,id);
        float originalDpi;using(var graphics=grids[0].CreateGraphics())originalDpi=graphics.DpiX;
        foreach(float factor in new[] {1f,1.25f,1.5f,2f})
        {
            // This is font-based stress coverage at the current display DPI. It
            // neither changes Windows scaling nor replaces real WPS DPI checks.
            var data=new CustomizationData();
            foreach(var target in targets.Where(t=>t.Control is DataGridView))
            {
                bool header=target.Caption;
                data.Formats[id+"::"+target.Id]=new FormatRule {Family=target.OriginalFont.Name,Size=(header?9f:10f)*factor,Style=header?1:0};
            }
            DialogFormatting.Apply(form,id,data);
            foreach(bool wide in new[] {false,true})
            {
                form.ClientSize=new Size(wide?Math.Max(1200,DialogCatalog.Find(id).Default.Width):Math.Max(320,form.MinimumSize.Width),Math.Max(240,DialogCatalog.Find(id).Default.Height));
                form.PerformLayout();
                foreach(var grid in grids)
                {
                    Ui.AutoFitColumns(grid);var sizes=form as DialogSizesForm;if(sizes!=null)sizes.FitColumns();
                    string label=id+" "+factor+"x font simulation / "+(wide?"wide":"narrow");
                    AssertGridHeaders(grid,label);
                    Check(grid.DefaultCellStyle.Font.Size==10f*factor,label+" uses the requested body font size");
                    foreach(DataGridViewColumn column in grid.Columns)
                    {
                        Check(column.HeaderCell.InheritedStyle.Font.Size==9f*factor,label+" / "+column.HeaderText+" uses the requested header font size");
                        Check(column.HeaderCell.InheritedStyle.Font.Bold,label+" / "+column.HeaderText+" retains default bold weight");
                    }
                    using(var graphics=grid.CreateGraphics())Check(graphics.DpiX==originalDpi,label+" preserves the actual display DPI");
                }
            }
        }
        DialogFormatting.Apply(form,id,new CustomizationData());
        form.ClientSize=DialogCatalog.Find(id).Default;form.PerformLayout();
    }
    // rule=R-27
    static void GridHeaderChecks()
    {
        int before=checks;DialogSizeStore.RootOverride=sandbox;DialogSizing.Initialize();
        string[] ids={
#if WPS_WORD
            "word.links",
#else
            "trace.precedents","trace.dependents","trace.errorsources","columns.filter","file.names","file.ocrprefixes","file.ocrfields","merge.headers","migration.data","elimination.receivable","elimination.revenuecost",
#endif
            "agent.dialogsizes"};
        foreach(string id in ids)
        using(var form=DialogPreviewFactory.Create(id,DialogCatalog.Find(id).Default))
        using(var parking=new Panel {Size=new Size(2000,1600)})
        {
            form.TopLevel=false;form.FormBorderStyle=FormBorderStyle.None;parking.Controls.Add(form);form.Show();form.PerformLayout();
            Check(!IsWindowVisible(form.Handle),id+" preview remains hidden and never opens a host");
            var grids=DialogFormatting.Walk(form).OfType<DataGridView>().ToArray();
            Check(grids.Length>0,id+" covers a real production grid");
            foreach(var grid in grids)
            {
                Ui.AutoFitColumns(grid);var sizes=form as DialogSizesForm;if(sizes!=null)sizes.FitColumns();
                AssertGridHeaders(grid,id+" default");
                Check(grid.Columns.Cast<DataGridViewColumn>().All(c=>c.HeaderCell.InheritedStyle.Font.Bold),id+" default headers are bold");
            }
            if(id=="file.names"||id=="file.ocrprefixes")RenderGridHeaders(form,id,"default-offline");
            GridHeaderFontScaleChecks(form,id,grids);
            var targets=DialogFormatting.Register(form,id);
            var data=new CustomizationData();
            // A body-only rule used to turn every header into a wrapping header.
            var body=targets.First(t=>t.Control is DataGridView&&!t.Caption);
            data.Formats[id+"::"+body.Id]=new FormatRule {Size=12,Family=body.OriginalFont.Name,Style=0};
            DialogFormatting.Apply(form,id,data);
            foreach(var grid in grids)AssertGridHeaders(grid,id+" body-only customization");
            var first=grids[0].Columns[0];float bodySize=grids[0].DefaultCellStyle.Font.Size;
            var header=targets.Single(t=>t.Control==grids[0]&&t.Id.Contains(".header.")&&t.OriginalText==first.HeaderText);
            data.Formats[id+"::"+header.Id]=new FormatRule {Text="完整表头中文 & long header 0123456789",Family=header.OriginalFont.Name,Size=36,Style=0};
            first.HeaderCell.Style.Padding=new Padding(17,3,19,5);
            DialogFormatting.Apply(form,id,data);
            foreach(var grid in grids)AssertGridHeaders(grid,id+" large explicit header");
            Check(!first.HeaderCell.InheritedStyle.Font.Bold,id+" explicit regular header font is respected");
            Check(grids[0].DefaultCellStyle.Font.Size==bodySize&&!grids[0].DefaultCellStyle.Font.Bold,id+" header format does not bold or enlarge body text");
            int floor=first.MinimumWidth;try{first.Width=2;}catch(ArgumentOutOfRangeException){}
            Check(first.Width>=floor,id+" manual narrowing cannot truncate the header");
            form.ClientSize=new Size(Math.Max(320,form.MinimumSize.Width),Math.Max(240,form.MinimumSize.Height));form.PerformLayout();
            foreach(var grid in grids){Ui.AutoFitColumns(grid);AssertGridHeaders(grid,id+" narrow window");}
            DialogFormatting.Apply(form,id,data);
            Check(!first.HeaderCell.InheritedStyle.Font.Bold,id+" repeated apply retains explicit regular weight");
            foreach(var grid in grids)AssertGridHeaders(grid,id+" repeated format application");
            if(id=="file.ocrprefixes")
            {
                var larger=new CustomizationData();foreach(var target in targets.Where(t=>t.Id.Contains(".header.")))larger.Formats[id+"::"+target.Id]=new FormatRule {Family=target.OriginalFont.Name,Size=18,Style=1};
                form.ClientSize=DialogCatalog.Find(id).Default;DialogFormatting.Apply(form,id,larger);
                foreach(var grid in grids)AssertGridHeaders(grid,id+" 18pt headers");RenderGridHeaders(form,id,"18pt-offline");
                form.ClientSize=new Size(1200,DialogCatalog.Find(id).Default.Height);form.PerformLayout();
                foreach(var grid in grids)AssertGridHeaders(grid,id+" 18pt wide window");RenderGridHeaders(form,id,"18pt-wide-offline");
            }
        }
        Console.WriteLine("Grid headers: "+(checks-before)+" PASS; production previews in hidden parents, no WPS; 1/1.25/1.5/2x font simulation, current display DPI unchanged");
    }
}
