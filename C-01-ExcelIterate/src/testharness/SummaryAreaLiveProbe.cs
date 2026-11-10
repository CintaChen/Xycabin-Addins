using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using Excel_Iterate;

// Explicit live test only. All documents are created and owned by ProbeSession.
internal static class SingleSummaryLiveProbe
{
    private static int checks;
    private static void Check(bool value,string name){if(!value)throw new Exception(name);checks++;Console.WriteLine("PASS "+name);}
    private static readonly BindingFlags Flags=BindingFlags.Instance|BindingFlags.NonPublic;
    private static void Set(object host,string name,object value){host.GetType().GetField(name,Flags).SetValue(host,value);}
    private static object Get(object host,string name){return host.GetType().GetField(name,Flags).GetValue(host);}
    private static void Call(object host,string name,params object[] args){try{host.GetType().GetMethod(name,Flags).Invoke(host,args);}catch(TargetInvocationException e){throw e.InnerException;}}
    private static void ProvidedHeaders(object application,ProbeSession session,string vectorsPath)
    {
        List<ProvidedHeaderVector> samples;
        using(var stream=System.IO.File.OpenRead(vectorsPath))samples=(List<ProvidedHeaderVector>)new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(List<ProvidedHeaderVector>)).ReadObject(stream);
        int number=0;
        foreach(var sample in samples)
        {
            number++;using(var bag=new ComBag())
            {
                dynamic app=application,books=bag.Own(app.Workbooks),book=books.Add();bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
                dynamic sheets=bag.Own(book.Worksheets),sheet=bag.Own(sheets.Item(1));sheet.Name="汇总";dynamic source=bag.Own(sheets.Add());source.Name=sample.sheet;
                foreach(var year in sample.years){dynamic cell=bag.Own(source.Range["N"+year.row]);cell.Value2=year.text;}
                foreach(var company in sample.companies){dynamic code=bag.Own(source.Range["O"+company.row]),name=bag.Own(source.Range["P"+company.row]);code.Value2=company.code;name.Value2=company.name;}
                int first=sample.companies[0].row,last=sample.companies[sample.companies.Count-1].row,width=sample.years.Count,oldGroups=Math.Min(3,sample.companies.Count);
                dynamic oldNames=bag.Own(sheet.Range["L2:"+CellAddress.Name(2,12+oldGroups*width-1)]);oldNames.FormulaArray="=INDEX('"+sample.sheet+"'!$O$"+first+":$O$"+(first+oldGroups-1)+",INT((COLUMN("+"L2:"+CellAddress.Name(2,12+oldGroups*width-1)+")-COLUMN(L2))/"+width+")+1)";oldNames.Calculate();string oldFormula=Convert.ToString(((dynamic)bag.Own(sheet.Range["L2"])).Formula);
                using(var form=new SingleSummaryForm(true))
                {
                    var view=(SummaryAreaView)form.Controls[0];view.Folder.Text=session.DirectoryPath;view.MatchHeaders.Checked=false;view.HeadersEnabled.Checked=true;view.FormulaHeaders.Checked=true;view.RepeatNames.Checked=true;view.ReplaceHeaderArrays.Checked=true;
                    using(var host=new SummaryAreaHost(application,form,view,null,delegate{return true;}))
                    {
                        Set(host,"book",(object)book);Set(host,"bookName",Convert.ToString(book.FullName));Set(host,"templateSheet","汇总");Set(host,"pasteSheet","汇总");Set(host,"leadSheet",sample.sheet);
                        ((List<SummaryAreaPlan>)Get(host,"areas")).Add(new SummaryAreaPlan{Sheet="汇总",Row=10,Column=7,Rows=1,Columns=width});Set(host,"paste",new SummaryAreaPlan{Row=10,Column=12,Rows=1,Columns=1});
                        Set(host,"lead",new SummaryAreaPlan{Sheet=sample.sheet,Row=first,Column=15,Rows=last-first+1,Columns=2});Set(host,"companyHeader",new SummaryAreaPlan{Sheet="汇总",Row=2,Column=12,Rows=1,Columns=1});Set(host,"yearSource",new SummaryAreaPlan{Sheet=sample.sheet,Row=sample.years[0].row,Column=14,Rows=width,Columns=1});Set(host,"yearTarget",new SummaryAreaPlan{Sheet="汇总",Row=7,Column=12,Rows=1,Columns=1});
                        Call(host,"Preview");Check((bool)Get(host,"ready"),"provided sample "+number+" actual N/O/P fields preflight");Call(host,"Start");for(int i=0;i<sample.companies.Count;i++)Call(host,"Tick",null,EventArgs.Empty);
                        bool correct=!(bool)Get(host,"failed")&&(bool)Get(host,"pending");for(int i=0;i<sample.companies.Count;i++)for(int c=0;c<width;c++){dynamic name=bag.Own(sheet.Range[CellAddress.Name(2,12+i*width+c)]),year=bag.Own(sheet.Range[CellAddress.Name(7,12+i*width+c)]);correct &= Convert.ToBoolean(name.HasFormula)&&Convert.ToString(name.Value2)==sample.companies[i].code&&Convert.ToBoolean(year.HasFormula)&&Convert.ToString(year.Value2)==sample.years[c].text;}
                        Check(correct,"provided sample "+number+" company order repetition and transposed years");Call(host,"Undo");Check(Convert.ToBoolean(oldNames.HasArray)&&Convert.ToString(((dynamic)bag.Own(sheet.Range["L2"])).Formula)==oldFormula,"provided sample "+number+" original array rollback");Check(Convert.ToString(book.Path)=="","provided sample "+number+" isolated workbook never saved");
                    }
                }
                // Exercise the new read-only column mode with the same supplied fields.
                oldNames.ClearContents();dynamic periodTitle=bag.Own(source.Range["N1"]);periodTitle.Value2="期间";
                int endColumn=7+sample.companies.Count*width-1;dynamic fixedNames=bag.Own(sheet.Range["G2:"+CellAddress.Name(2,endColumn)]);var nameRefs=new List<string>();foreach(var item in sample.companies)nameRefs.Add("'"+sample.sheet+"'!$O$"+item.row);
                fixedNames.FormulaArray="=CHOOSE(INT((COLUMN(G2:"+CellAddress.Name(2,endColumn)+")-COLUMN(G2))/"+width+")+1,"+string.Join(",",nameRefs.ToArray())+")";fixedNames.Calculate();string fixedBefore=Convert.ToString(((dynamic)bag.Own(sheet.Range["G2"])).Formula);
                for(int group=0;group<sample.companies.Count;group++)
                {
                    dynamic yearCells=bag.Own(sheet.Range[CellAddress.Name(7,7+group*width)+":"+CellAddress.Name(7,7+(group+1)*width-1)]);var row=new object[1,width];for(int c=0;c<width;c++)row[0,c]=sample.years[group%2==0?c:width-1-c].text;yearCells.Value2=row;
                    string filename=sample.companies[group].name+".xlsx",file=System.IO.Path.Combine(session.DirectoryPath,filename);if(!System.IO.File.Exists(file))System.IO.File.Copy(System.IO.Path.Combine(session.DirectoryPath,"！01-示例（集团）.xlsx"),session.Output(filename));
                }
                for(int c=0;c<width;c++){dynamic cell=bag.Own(sheet.Range[CellAddress.Name(10,7+c)]);cell.Value2="BS!$A$"+(c+1);}
                using(var form=new SingleSummaryForm(true))
                {
                    var view=(SummaryAreaView)form.Controls[0];view.Folder.Text=session.DirectoryPath;view.MatchHeaders.Checked=true;
                    using(var host=new SummaryAreaHost(application,form,view,null,delegate{return true;}))
                    {
                        Set(host,"book",(object)book);Set(host,"bookName",Convert.ToString(book.FullName));Set(host,"templateSheet","汇总");Set(host,"leadSheet",sample.sheet);Set(host,"lead",new SummaryAreaPlan{Sheet=sample.sheet,Row=first,Column=15,Rows=last-first+1,Columns=2});((List<SummaryAreaPlan>)Get(host,"areas")).Add(new SummaryAreaPlan{Sheet="汇总",Row=10,Column=7,Rows=1,Columns=width});
                        if(sample.companies.Count==1)
                        {bool rejected=false;try{Call(host,"Preview");}catch(InvalidOperationException){rejected=true;}Check(rejected,"provided fixed sample "+number+" skips sole already-compiled company");}
                        else
                        {
                            Call(host,"Preview");Check((bool)Get(host,"ready"),"provided fixed sample "+number+" existing header discovery");Call(host,"Start");int tasks=((List<SummaryAreaPlan>)Get(host,"targets")).Count;for(int i=0;i<tasks;i++)Call(host,"Tick",null,EventArgs.Empty);
                            bool good=!(bool)Get(host,"failed");for(int group=1;group<sample.companies.Count;group++)for(int c=0;c<width;c++){dynamic cell=bag.Own(sheet.Range[CellAddress.Name(10,7+group*width+c)]);string formula=Convert.ToString(cell.Formula);int index=group%2==0?c:width-1-c;good &= formula.Contains("["+sample.companies[group].name+".xlsx]BS")&&formula.EndsWith("!$A$"+(index+1),StringComparison.Ordinal);}
                            Check(good,"provided fixed sample "+number+" exact company and year column links");Call(host,"Undo");Check(Convert.ToBoolean(fixedNames.HasArray)&&Convert.ToString(((dynamic)bag.Own(sheet.Range["G2"])).Formula)==fixedBefore,"provided fixed sample "+number+" data undo preserves header declaration");
                        }
                    }
                }
                book.Close(false);closed=true;
            }
        }
    }
    private static void FixedHeaders(object application,ProbeSession session)
    {
        using(var bag=new ComBag())
        {
            dynamic app=application,books=bag.Own(app.Workbooks),book=books.Add();bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
            dynamic sheets=bag.Own(book.Worksheets),target=bag.Own(sheets.Item(1));target.Name="列名汇总";dynamic lead=bag.Own(sheets.Add());lead.Name="Lead";
            dynamic leadNames=bag.Own(lead.Range["O3:P5"]);leadNames.Value2=new object[,]{{"甲简称","！01-示例（集团）"},{"乙简称","！02-示例乙"},{"丙简称","第三家公司"}};
            dynamic heading=bag.Own(lead.Range["N1"]);heading.Value2="期间";dynamic dates=bag.Own(lead.Range["N2:N6"]);dates.Value2=new object[,]{{"2027年度"},{"2026年度"},{"2025年度"},{"2024年度"},{"2023年度"}};
            dynamic names=bag.Own(target.Range["G2:U2"]);names.FormulaArray="=INDEX(Lead!$O$3:$O$5,INT((COLUMN(G2:U2)-COLUMN(G2))/5)+1)";names.Calculate();string nameBefore=Convert.ToString(((dynamic)bag.Own(target.Range["G2"])).Formula);
            foreach(string address in new[]{"G7:K7","Q7:U7"}){dynamic year=bag.Own(target.Range[address]);year.FormulaArray="=TRANSPOSE(Lead!$N$2:$N$6)";year.Calculate();}
            dynamic reverse=bag.Own(target.Range["L7:P7"]);reverse.Value2=new object[,]{{"2023年度","2024年度","2025年度","2026年度","2027年度"}};
            string yearBefore=Convert.ToString(((dynamic)bag.Own(target.Range["G7"])).Formula);
            for(int c=0;c<5;c++)foreach(int r in new[]{3,4,8,9,10,11,12}){dynamic cell=bag.Own(target.Range[CellAddress.Name(r,7+c)]);cell.Value2="BS!$A$"+(c+1);}
            dynamic native=bag.Own(target.Range["G8"]);native.Formula="=ROUND(G9-G10,2)";dynamic old=bag.Own(target.Range["P3"]);old.Value2="原值";
            dynamic gap=bag.Own(target.Range["L5"]);gap.Value2="未选空隙";bag.Own(gap.AddComment("原批注"));
            using(var form=new SingleSummaryForm(true))
            {
                var view=(SummaryAreaView)form.Controls[0];view.Folder.Text=session.DirectoryPath;view.MatchHeaders.Checked=true;
                using(var host=new SummaryAreaHost(application,form,view,null,delegate{return true;}))
                {
                    Set(host,"book",(object)book);Set(host,"bookName",Convert.ToString(book.FullName));Set(host,"templateSheet","列名汇总");Set(host,"leadSheet","Lead");
                    var selected=(List<SummaryAreaPlan>)Get(host,"areas");selected.Add(new SummaryAreaPlan{Sheet="列名汇总",Row=2,Column=7,Rows=3,Columns=5});selected.Add(new SummaryAreaPlan{Sheet="列名汇总",Row=7,Column=7,Rows=6,Columns=5});Set(host,"lead",new SummaryAreaPlan{Sheet="Lead",Row=3,Column=15,Rows=3,Columns=2});
                    int before=Convert.ToInt32(sheets.Count);Call(host,"Preview");Check((bool)Get(host,"ready")&&Convert.ToInt32(sheets.Count)==before,"fixed header preflight is read-only and requires no paste/header inputs");
                    Check(((List<SummaryCompany>)Get(host,"companies")).Find(c=>c.Row==3).Selected==false&&Convert.ToString(view.Companies.Rows[0].Cells["status"].Value)=="已编制，跳过","existing source company is skipped with explicit status");
                    var work=(List<SummaryMappedWork>)Get(host,"mappedWork");Check(work.TrueForAll(w=>w.Target.Row!=2&&w.Target.Row!=7),"company and year header rows excluded from all writes");
                    Call(host,"Start");int count=((List<SummaryAreaPlan>)Get(host,"targets")).Count;for(int i=0;i<count;i++)Call(host,"Tick",null,EventArgs.Empty);Check(!(bool)Get(host,"failed")&&(bool)Get(host,"pending"),"fixed column execution completes with reused array headers");
                    dynamic p3=bag.Own(target.Range["P3"]),l3=bag.Own(target.Range["L3"]),p8=bag.Own(target.Range["P8"]),q3=bag.Own(target.Range["Q3"]);
                    Check(Convert.ToString(p3.Formula).Contains("[！02-示例乙.xlsx]BS'!$A$1")&&Convert.ToString(l3.Formula).Contains("!$A$5"),"links follow 2027/2023 labels when target years are reversed");
                    Check(Convert.ToString(p8.Formula)=="=ROUND(P9-P10,2)","native same-year calculation follows matched target column");
                    Check(Convert.ToString(q3.Formula).Contains("[第三家公司.xlsx]BS'!$A$1"),"next company follows its own existing header group");
                    Check(Convert.ToString(((dynamic)bag.Own(target.Range["G2"])).Formula)==nameBefore&&Convert.ToString(((dynamic)bag.Own(target.Range["G7"])).Formula)==yearBefore&&Convert.ToBoolean(names.HasArray),"original company and year array declarations remain unchanged");
                    Check(Convert.ToString(gap.Value2)=="未选空隙"&&Convert.ToString(((dynamic)bag.Own(gap.Comment)).Text())=="原批注","Ctrl gaps retain value and note");
                    Call(host,"Undo");Check(Convert.ToString(old.Value2)=="原值"&&Convert.ToString(((dynamic)bag.Own(target.Range["G2"])).Formula)==nameBefore,"undo restores data without rewriting headers");
                    var preset=host.CapturePreset();host.ApplyPreset(preset);Check(view.MatchHeaders.Checked&&!(bool)Get(host,"ready"),"column mode preset accepts absent paste/header inputs and invalidates plan");
                    Call(host,"Preview");dynamic changed=bag.Own(target.Range["L7"]);changed.Value2="2024年度";bool stopped=false;try{Call(host,"Start");}catch(InvalidOperationException){stopped=true;}Check(stopped,"header value edit invalidates the preflight");
                    stopped=false;try{Call(host,"Preview");}catch(InvalidOperationException){stopped=true;}Check(stopped,"duplicate or missing company-year labels block execution");changed.Value2="2023年度";
                    reverse.ClearContents();view.FillMissingYears.Checked=true;view.MissingYearSequence.Text="2023年度、2024年度、2025年度、2026年度、2027年度";
                    Call(host,"Preview");Check((bool)Get(host,"ready")&&((List<SummaryAreaPlan>)Get(host,"headerTargets")).Count==5&&string.IsNullOrEmpty((string)Convert.ToString(changed.Value2)),"missing years preflight plans five blanks without writing; ready="+Get(host,"ready")+"; count="+((List<SummaryAreaPlan>)Get(host,"headerTargets")).Count+"; L7="+Convert.ToString(changed.Value2));
                    Call(host,"Start");count=((List<SummaryAreaPlan>)Get(host,"targets")).Count;for(int i=0;i<count;i++)Call(host,"Tick",null,EventArgs.Empty);
                    Check(!(bool)Get(host,"failed")&&(bool)Get(host,"pending")&&Convert.ToString(changed.Value2)=="2023年度"&&Convert.ToString(((dynamic)bag.Own(target.Range["P7"])).Value2)=="2027年度","custom reversed years filled and verified");
                    Check(Convert.ToString(p3.Formula).Contains("!$A$1")&&Convert.ToString(l3.Formula).Contains("!$A$5")&&Convert.ToString(((dynamic)bag.Own(target.Range["G2"])).Formula)==nameBefore&&Convert.ToString(((dynamic)bag.Own(target.Range["G7"])).Formula)==yearBefore,"new year plan drives correct links and preserves existing arrays");
                    Call(host,"Undo");Check(string.IsNullOrEmpty((string)Convert.ToString(changed.Value2))&&string.IsNullOrEmpty((string)Convert.ToString(((dynamic)bag.Own(target.Range["P7"])).Value2))&&Convert.ToString(old.Value2)=="原值","undo restores missing headers and data together");
                    preset=host.CapturePreset();host.ApplyPreset(preset);Check(view.FillMissingYears.Checked&&view.MissingYearSequence.Text.StartsWith("2023年度"),"missing year settings persist and restore");
                    changed.Value2="2027年度";stopped=false;try{Call(host,"Preview");}catch(InvalidOperationException){stopped=true;}Check(stopped,"partial existing year conflicts are blocked before writing");
                    changed.Formula="=\"\"";stopped=false;try{Call(host,"Preview");}catch(InvalidOperationException){stopped=true;}Check(stopped,"blank formula cannot be overwritten by missing year fill");changed.ClearContents();
                    dynamic merged=bag.Own(target.Range["L7:M7"]);merged.Merge();stopped=false;try{Call(host,"Preview");}catch(InvalidOperationException){stopped=true;}Check(stopped,"merged missing year cells block execution");merged.UnMerge();
                    view.FillMissingYears.Checked=false;stopped=false;try{Call(host,"Preview");}catch(InvalidOperationException){stopped=true;}Check(stopped,"disabled year fill retains missing header protection");
                    Check(Convert.ToString(book.Path)=="","fixed mapping never saves workbook");
                }
            }
            book.Close(false);closed=true;
        }
    }
    [STAThread] private static int Main(string[] args)
    {
        if (Array.IndexOf(args, "--timing") >= 0) return SummaryAreaSpeedChecks.Run();
        using(var session=new ProbeSession("SummaryAreaLive"))
        {
            try
            {
                Type type=Type.GetTypeFromProgID("Excel.Application",false);
                if(type==null){Console.WriteLine("SKIP: Excel.Application not registered");return session.Finish(2);}
                dynamic app=session.Take(Activator.CreateInstance(type));
                Console.WriteLine("HOST: "+Convert.ToString(app.Name)+"; version="+Convert.ToString(app.Version));
                if(!session.Override(()=>app.DisplayAlerts,v=>app.DisplayAlerts=v,false,"DisplayAlerts") ||
                   !session.Override(()=>app.EnableEvents,v=>app.EnableEvents=v,false,"EnableEvents") ||
                   !session.Override(()=>app.ScreenUpdating,v=>app.ScreenUpdating=v,false,"ScreenUpdating"))throw new Exception("Cannot establish isolated host state");
                using(var bag=new ComBag())
                {
                    dynamic books=bag.Own(app.Workbooks);
                    foreach(string name in new[]{"！01-示例（集团）","！02-示例乙","公共参数","第三家公司"})
                    {
                        dynamic source=books.Add(); bool closed=false;
                        session.Own((object)source,delegate{if(!closed)source.Close(false);});
                        using(var refs=new ComBag()) { dynamic sheets=refs.Own(source.Worksheets), sheet=refs.Own(sheets.Item(1));sheet.Name="BS";dynamic cell=refs.Own(sheet.Range["A1"]);cell.Value2=25; }
                        source.SaveAs(session.Output(name+".xlsx"),51);source.Close(false);closed=true;
                    }
                    if(Array.IndexOf(args,"--fixed-only")>=0){FixedHeaders((object)app,session);Console.WriteLine("PASS: "+checks+" fixed header checks");return session.Finish(0);}
                    dynamic book=books.Add();session.Own((object)book,delegate{book.Close(false);});
                    dynamic worksheets=bag.Own(book.Worksheets), target=bag.Own(worksheets.Item(1));target.Name="汇总";
                    dynamic lead=bag.Own(worksheets.Add());lead.Name="Lead";
                    dynamic leadRange=bag.Own(lead.Range["A1:B2"]);leadRange.Value2=new object[,]{{"01","！01-示例（集团）"},{"02","！02-示例乙"}};
                    foreach(string address in new[]{"L3","L6"}){dynamic cell=bag.Own(target.Range[address]);cell.Value2="BS!$A$1";}
                    dynamic native=bag.Own(target.Range["M3"]);native.Formula="=ROUND(L3-L3,2)";
                    dynamic blank=bag.Own(target.Range["M6"]), fill=bag.Own(blank.Interior);fill.Color=65535;
                    dynamic destination=bag.Own(target.Range["Q3:T6"]);destination.Value2="原值";
                    dynamic gap=bag.Own(target.Range["Q4"]);gap.Value2="未选空隙";bag.Own(gap.AddComment("空隙批注"));dynamic gapFill=bag.Own(gap.Interior);gapFill.Color=12345;
                    dynamic old=bag.Own(target.Range["Q3"]);old.Formula="='汇总'!A10+7";string oldFormula=Convert.ToString(old.Formula);
                    dynamic note=bag.Own(target.Range["R6"]);bag.Own(note.AddComment("原目标批注"));
                    int before=Convert.ToInt32(worksheets.Count);
                    using(var form=new SingleSummaryForm(true))
                    {
                        var view=(SummaryAreaView)form.Controls[0];view.Folder.Text=session.DirectoryPath;view.MatchHeaders.Checked=false;view.FormulaHeaders.Checked=false;view.RepeatNames.Checked=false;
                        using(var host=new SummaryAreaHost((object)app,form,view,null,delegate{return true;}))
                        {
                            Set(host,"book",(object)book);Set(host,"bookName",Convert.ToString(book.FullName));Set(host,"templateSheet","汇总");Set(host,"pasteSheet","汇总");Set(host,"leadSheet","Lead");
                            var areas=(List<SummaryAreaPlan>)Get(host,"areas");areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=3,Column=12,Rows=1,Columns=2,Address="L3:M3"});areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=6,Column=12,Rows=1,Columns=2,Address="L6:M6"});
                            Set(host,"paste",new SummaryAreaPlan{Row=3,Column=17,Rows=1,Columns=1});Set(host,"lead",new SummaryAreaPlan{Sheet="Lead",Row=1,Column=1,Rows=2,Columns=2,Address="A1:B2"});
                            Call(host,"Preview");Check((bool)Get(host,"ready"),"read-only preflight succeeds");Check(Convert.ToInt32(worksheets.Count)==before,"preflight creates no backup sheet");Check(Convert.ToString(old.Formula)==oldFormula,"preflight leaves target formula untouched");
                            int calculationBefore=Convert.ToInt32(app.Calculation);Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);
                            Check(Convert.ToInt32(app.Calculation)==calculationBefore,"batch restores original calculation mode");Check(Convert.ToInt32(Get(host,"bulkWrites"))==4,"four company Areas use four matrix formula submissions");
                            Check((bool)Get(host,"pending")&&!(bool)Get(host,"running")&&!(bool)Get(host,"failed"),"execution reaches pending confirmation");
                            dynamic q3=bag.Own(target.Range["Q3"]),s3=bag.Own(target.Range["S3"]),r3=bag.Own(target.Range["R3"]),r6=bag.Own(target.Range["R6"]);
                            Check(Convert.ToBoolean(q3.HasFormula)&&Convert.ToString(q3.Formula).Contains("[！01-示例（集团）.xlsx]BS"),"first company real external formula");
                            Check(Convert.ToString(s3.Formula).Contains("[！02-示例乙.xlsx]BS"),"second company from original template");
                            Check(Convert.ToString(r3.Formula)=="=ROUND(Q3-Q3,2)","native relative formula copied horizontally");
                            Check(r6.Value2==null,"selected blank replaces old target content");Check(Convert.ToInt32(((dynamic)bag.Own(r6.Interior)).Color)==65535,"selected blank retains template formatting");
                            Check(Convert.ToString(gap.Value2)=="未选空隙"&&Convert.ToInt32(gapFill.Color)==12345,"unselected gap value and format unchanged");Check(Convert.ToString(((dynamic)bag.Own(gap.Comment)).Text())=="空隙批注","unselected gap comment unchanged");
                            Check(Convert.ToString(((dynamic)bag.Own(target.Range["L3"])).Value2)=="BS!$A$1","original template remains text");
                            Call(host,"Undo");Check(Convert.ToString(q3.Formula)==oldFormula,"rollback restores exact explicit self-sheet formula");Check(Convert.ToString(((dynamic)bag.Own(r6.Comment)).Text())=="原目标批注","rollback restores target comment");Check(Convert.ToString(r6.Value2)=="原值","rollback restores selected blank target");
                            Call(host,"Preview");dynamic change=bag.Own(target.Range["L3"]);change.Value2="BS!$A$2";bool rejected=false;try{Call(host,"Start");}catch(InvalidOperationException){rejected=true;}Check(rejected,"template edit invalidates preflight without relying on events");change.Value2="BS!$A$1";
                            Call(host,"Preview");Call(host,"Start");Call(host,"Tick",null,EventArgs.Empty);target.Protect();Call(host,"Tick",null,EventArgs.Empty);
                            Check((bool)Get(host,"failed")&&(bool)Get(host,"pending"),"injected protection failure keeps recovery session");target.Unprotect();Call(host,"Undo");Check(!(bool)Get(host,"pending")&&Convert.ToString(q3.Formula)==oldFormula,"retry undo recovers failed batch");
                            Call(host,"Preview");Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);Call(host,"Confirm");Check(!(bool)Get(host,"pending"),"confirmation ends rollback session");Check(Convert.ToString(book.Path)=="","target workbook was never saved");
                            target.Activate();lead.Name="lEaD";Check(host.ActivateLeadSheet(),"Lead jump recognizes case-insensitive sheet name");
                            dynamic active=bag.Own(app.ActiveSheet);Check(Convert.ToString(active.Name)=="lEaD","Lead jump activates actual Lead worksheet");
                            lead.Name="公司清单";target.Activate();Check(!host.ActivateLeadSheet(),"missing Lead keeps manual worksheet selection available");lead.Name="Lead";
                            foreach(var pair in new[]{new[]{"A7","01"},new[]{"D7","！01-示例（集团）"},new[]{"A4","02"},new[]{"D4","！02-示例乙"},new[]{"A5","未选行"},new[]{"D5","未选行不得读取"}}){dynamic cell=bag.Own(lead.Range[pair[0]]);cell.Value2=pair[1];}
                            lead.Activate();dynamic union=bag.Own(lead.Range["A7,D7,A4,D4"]);union.Select();
                            dynamic unionAreas=bag.Own(union.Areas),firstArea=bag.Own(unionAreas.Item(1));int expectedFirst=Convert.ToInt32(firstArea.Row);
                            Call(host,"Pick","lead",false);Check(((List<SummaryAreaPlan>)Get(host,"leadAreas")).Count==4,"actual Excel Selection.Areas captured for disjoint Lead fields");
                            Call(host,"Preview");Check((bool)Get(host,"ready")&&view.Companies.Rows.Count==2,"disjoint Lead preflight ignores unselected middle rows");
                            Check(Convert.ToInt32(view.Companies.Rows[0].Cells["row"].Value)==expectedFirst,"disjoint Lead retains Excel Areas order");
                            Check(Convert.ToString(view.Companies.Rows[0].Cells["company"].Value)==(expectedFirst==7?"！01-示例（集团）":"！02-示例乙"),"nonadjacent code and name columns paired by same row");
                            view.Folder.Text="\""+session.DirectoryPath+"\"";Check(!(bool)Get(host,"ready"),"manual folder edit invalidates prior plan");Call(host,"Preview");Check((bool)Get(host,"ready"),"manually pasted quoted folder passes preflight");
                            var saved=host.CapturePreset();string presetPath=session.Output("区域预设.xybatch.json");BatchPresetStore.Write(presetPath,saved,false);session.TrackFile(presetPath+".lock");saved=BatchPresetStore.Read(presetPath,"summary");
                            Check(!saved.Values.ContainsKey("companies")&&saved.Values["leadAreas"].Split(';').Length==4,"saved presets contain settings and all Areas, not company data");
                            Check((bool)typeof(SummaryAreaHost).GetMethod("ContextAvailable",Flags).Invoke(host,new object[]{saved.Context}),"unsaved open workbook context can be reused without file relocation");
                            view.Folder.Text="待替换路径";var invalid=BatchPresetStore.Decode(BatchPresetStore.Encode(saved),"summary");invalid.Values["templateSheet"]="不存在的工作表";
                            bool failedApply=false;try{host.ApplyPreset(invalid);}catch{failedApply=true;}Check(failedApply&&view.Folder.Text=="待替换路径","invalid preset retains prior UI and workbook binding");
                            var legacy=BatchPresetStore.Decode(BatchPresetStore.Encode(saved),"summary");legacy.Values.Remove("mode");string legacyPath=session.Output("旧整列预设.xybatch.json");BatchPresetStore.Write(legacyPath,legacy,false);session.TrackFile(legacyPath+".lock");string legacyBytes=Convert.ToBase64String(System.IO.File.ReadAllBytes(legacyPath));
                            failedApply=false;try{host.ApplyPreset(legacy);}catch(InvalidOperationException){failedApply=true;}Check(failedApply&&legacyBytes==Convert.ToBase64String(System.IO.File.ReadAllBytes(legacyPath)),"legacy preset remains unchanged and is not guessed into area mode");
                            int sheetsBefore=Convert.ToInt32(worksheets.Count);host.ApplyPreset(saved);Check(!(bool)Get(host,"ready")&&view.Companies.Rows.Count==0&&view.Folder.Text==saved.Values["projectFolder"],"apply restores parameters and requires a fresh company preflight");
                            Check(Convert.ToInt32(worksheets.Count)==sheetsBefore&&Convert.ToString(book.Path)=="","applying preset creates no backup and never saves workbook");
                            Call(host,"Preview");Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);failedApply=false;try{host.ApplyPreset(saved);}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"pending rollback blocks preset application");Call(host,"Undo");
                            Call(host,"Preview");dynamic formatCell=bag.Own(target.Range["L3"]),formatFont=bag.Own(formatCell.Font);object priorColor=formatFont.Color;formatFont.Color=12345;failedApply=false;try{Call(host,"Start");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"native block fingerprint rejects a font-only change before execution");formatFont.Color=priorColor;
                            dynamic priorNote=bag.Own(formatCell.AddComment("状态校验批注"));Call(host,"Preview");priorNote.Text("预检后修改批注");failedApply=false;try{Call(host,"Start");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"native block fingerprint rejects comment-only change before execution");priorNote.Text("状态校验批注");
                            dynamic computedName=bag.Own(lead.Range["D7"]),nameSource=bag.Own(lead.Range["Z7"]);nameSource.Value2="！01-示例（集团）";computedName.Formula="=Z7";computedName.Calculate();Call(host,"Preview");nameSource.Value2="！02-示例乙";computedName.Calculate();failedApply=false;try{Call(host,"Start");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"computed Lead company name change invalidates plan even with unchanged formula text");computedName.Value2="！01-示例（集团）";nameSource.Clear();
                            Check(!view.HeadersEnabled.Checked,"headers default off");view.FormulaHeaders.Checked=false;view.RepeatNames.Checked=false;
                            dynamic yearSource=bag.Own(lead.Range["G1:H1"]);yearSource.Value2=new object[,]{{46388,"2026年度"}};
                            dynamic dated=bag.Own(lead.Range["G1"]);dated.NumberFormat="yyyy年";string displayedYear=Convert.ToString(dated.Text);
                            dynamic firstName=bag.Own(target.Range["Q2:R2"]),secondName=bag.Own(target.Range["S2:T2"]);firstName.Merge();secondName.Merge();
                            dynamic firstNameCell=bag.Own(target.Range["Q2"]);firstNameCell.Formula="='汇总'!A10+12";string nameBefore=Convert.ToString(firstNameCell.Formula);
                            dynamic oldYears=bag.Own(target.Range["Q5:T5"]);oldYears.Value2="旧年度";oldYears.NumberFormat="0.00";
                            target.Activate();firstName.Select();Call(host,"Pick","companyHeader",false);Check(((SummaryAreaPlan)Get(host,"companyHeader")).Columns==1,"merged selection normalizes to name anchor");
                            lead.Activate();yearSource.Select();Call(host,"Pick","yearSource",false);target.Activate();dynamic yearAnchor=bag.Own(target.Range["Q5"]);yearAnchor.Select();Call(host,"Pick","yearTarget",false);view.HeadersEnabled.Checked=true;
                            Call(host,"Preview");Check((bool)Get(host,"ready")&&((List<SummaryAreaPlan>)Get(host,"headerTargets")).Count==6,"merged names and years pass read-only preflight");Check(Convert.ToString(firstNameCell.Formula)==nameBefore,"header preflight leaves formula untouched");
                            string expectedName1=Convert.ToString(view.Companies.Rows[0].Cells["code"].Value),expectedName2=Convert.ToString(view.Companies.Rows[1].Cells["code"].Value);
                            var headerPreset=host.CapturePreset();host.ApplyPreset(headerPreset);Check(view.HeadersEnabled.Checked&&Get(host,"yearSource")!=null,"header presets round-trip");
                            Call(host,"Preview");Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);
                            Check((bool)Get(host,"pending")&&!(bool)Get(host,"failed"),"header batch reaches pending");
                            Check(Convert.ToString(firstNameCell.Value2)==expectedName1&&Convert.ToString(((dynamic)bag.Own(target.Range["S2"])).Value2)==expectedName2,"Lead first column maps to company groups");
                            Check(Convert.ToString(yearAnchor.Value2)==displayedYear&&Convert.ToString(((dynamic)bag.Own(target.Range["S5"])).Value2)==displayedYear,"date source repeats display text instead of serial");
                            Check(Convert.ToString(((dynamic)bag.Own(target.Range["R5"])).Value2)=="2026年度"&&Convert.ToString(((dynamic)bag.Own(target.Range["T5"])).Value2)=="2026年度","year source order repeats");
                            Check(Convert.ToBoolean(firstName.MergeCells)&&Convert.ToString(gap.Value2)=="未选空隙","legal merges and unselected gaps unchanged");
                            Call(host,"Undo");Check(Convert.ToString(firstNameCell.Formula)==nameBefore&&Convert.ToString(yearAnchor.Value2)=="旧年度"&&Convert.ToString(yearAnchor.NumberFormat)=="0.00","undo restores header formula values and format");
                            Check(Convert.ToString(q3.Formula).Contains(".xlsx]BS"),"header undo restores prior link batch");
                            Call(host,"Preview");dated.NumberFormat="yyyy";failedApply=false;try{Call(host,"Start");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"year display change invalidates preflight");dated.NumberFormat="yyyy年";
                            Set(host,"yearSource",new SummaryAreaPlan{Sheet="Lead",Row=1,Column=7,Rows=1,Columns=1});failedApply=false;try{Call(host,"Preview");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"year count mismatch blocked");Set(host,"yearSource",new SummaryAreaPlan{Sheet="Lead",Row=1,Column=7,Rows=1,Columns=2});
                            Set(host,"yearTarget",new SummaryAreaPlan{Sheet="汇总",Row=3,Column=17,Rows=1,Columns=1});failedApply=false;try{Call(host,"Preview");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"header and copy conflict blocked");Set(host,"yearTarget",new SummaryAreaPlan{Sheet="汇总",Row=5,Column=17,Rows=1,Columns=1});
                            firstName.UnMerge();secondName.UnMerge();dynamic illegal=bag.Own(target.Range["Q2:S2"]);illegal.Merge();failedApply=false;try{Call(host,"Preview");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"cross-company merged name blocked");illegal.UnMerge();
                            host.ApplyPreset(saved);Check(!view.HeadersEnabled.Checked&&Get(host,"companyHeader")==null,"old preset disables optional headers");Call(host,"Preview");Check((bool)Get(host,"ready"),"V1 remains usable after header validation failures");
                            // Three companies, width five, two disjoint Areas; selected gaps stay untouched.
                            areas.Clear();areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=10,Column=12,Rows=1,Columns=5});areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=14,Column=12,Rows=1,Columns=5});
                            dynamic fiveTemplate=bag.Own(target.Range["L10:P10,L14:P14"]);fiveTemplate.Value2="BS!$A$1";
                            dynamic threeLead=bag.Own(lead.Range["A20:B22"]);threeLead.Value2=new object[,]{{"甲简称","！01-示例（集团）"},{"乙简称","！02-示例乙"},{"丙简称","第三家公司"}};
                            ((List<SummaryAreaPlan>)Get(host,"leadAreas")).Clear();Set(host,"lead",new SummaryAreaPlan{Sheet="Lead",Row=20,Column=1,Rows=3,Columns=2});view.Companies.Rows.Clear();Set(host,"paste",new SummaryAreaPlan{Row=10,Column=17,Rows=1,Columns=1});
                            dynamic fiveYears=bag.Own(lead.Range["G2:K2"]);fiveYears.Value2=new object[,]{{"2027年度","2026年度","2025年度","2024年度","2023年度"}};
                            Set(host,"companyHeader",new SummaryAreaPlan{Sheet="汇总",Row=9,Column=17,Rows=1,Columns=1});Set(host,"yearSource",new SummaryAreaPlan{Sheet="Lead",Row=2,Column=7,Rows=1,Columns=5});Set(host,"yearTarget",new SummaryAreaPlan{Sheet="汇总",Row=12,Column=17,Rows=1,Columns=1});view.HeadersEnabled.Checked=true;
                            dynamic untouched=bag.Own(target.Range["R11"]);untouched.Value2="未授权间隙";
                            Call(host,"Preview");Check(((List<SummaryAreaPlan>)Get(host,"headerTargets")).Count==18,"three companies width five plan has eighteen header tasks");Call(host,"Start");for(int i=0;i<6;i++)Call(host,"Tick",null,EventArgs.Empty);
                            Check(!(bool)Get(host,"failed")&&(bool)Get(host,"pending"),"three-company five-column multi-Area execution succeeds");
                            bool mapping=true;for(int i=0;i<3;i++){dynamic name=bag.Own(target.Cells[9,17+i*5]);mapping &= Convert.ToString(name.Value2)==new[]{"甲简称","乙简称","丙简称"}[i];for(int c=0;c<5;c++){dynamic year=bag.Own(target.Cells[12,17+i*5+c]);mapping &= Convert.ToString(year.Value2)==(2027-c)+"年度";}}
                            Check(mapping,"Q V AA company names and all fifteen years map correctly");Check(Convert.ToString(untouched.Value2)=="未授权间隙","three-company unselected gap remains untouched");
                            Call(host,"Undo");Check(((dynamic)bag.Own(target.Range["Q9"])).Value2==null&&((dynamic)bag.Own(target.Range["Q12"])).Value2==null&&((dynamic)bag.Own(target.Range["Q10"])).Value2==null,"three-company rollback restores blank headers and link targets");
                            Call(host,"Preview");Call(host,"Start");Call(host,"Tick",null,EventArgs.Empty);typeof(SummaryAreaHost).GetMethod("WriteHeaders",Flags).Invoke(host,new object[]{target,bag});
                            target.Protect();for(int i=0;i<2;i++)Call(host,"Tick",null,EventArgs.Empty);target.Unprotect();Call(host,"Undo");Check(((dynamic)bag.Own(target.Range["Q9"])).Value2==null&&((dynamic)bag.Own(target.Range["Q12"])).Value2==null,"failed batch retry undo includes headers already touched");
                            view.Companies.Rows[1].Cells["selected"].Value=false;Call(host,"Preview");Check(((List<SummaryAreaPlan>)Get(host,"headerTargets")).Count==12,"company filtering rebuilds header plan with two groups");Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);
                            Check(Convert.ToString(((dynamic)bag.Own(target.Range["V9"])).Value2)=="丙简称"&&((dynamic)bag.Own(target.Range["AA9"])).Value2==null,"unchecked middle company does not shift name and link pairing");Call(host,"Undo");
                            // Existing compiled Area: company external + unrelated public external + constants.
                            dynamic compiledSource=bag.Own(target.Range["L10"]),publicSource=bag.Own(target.Range["M10"]),constant=bag.Own(target.Range["N10"]),literalText=bag.Own(target.Range["O10"]);
                            compiledSource.Formula=SummaryTextReference.External(session.DirectoryPath,"！01-示例（集团）","BS!A1")+"+A1";
                            publicSource.Formula=SummaryTextReference.External(session.DirectoryPath,"公共参数","BS!$A$1");constant.Value2=123;literalText.Value2="说明! 请检查";dynamic formulaText=bag.Own(target.Range["P10"]);formulaText.NumberFormat="@";formulaText.Value2="=1+1";
                            bool reviewBlocked=false;try{Call(host,"Preview");}catch(InvalidOperationException){reviewBlocked=true;}
                            Check(reviewBlocked&&view.ReferenceSources.Rows.Count==2,"unrelated sources are reported and block first preflight");
                            foreach(DataGridViewRow reviewRow in view.ReferenceSources.Rows)reviewRow.Cells["decision"].Value=Convert.ToString(reviewRow.Cells["source"].Value).Contains("公共参数")?"保留原引用":"参与公司替换";
                            Call(host,"Preview");Check((bool)Get(host,"ready"),"explicit source decisions permit fresh preflight");string publicBefore=Convert.ToString(publicSource.Formula),compiledBefore=Convert.ToString(compiledSource.Formula);
                            Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);Check(!(bool)Get(host,"failed"),"existing compiled Area executes");
                            dynamic secondLink=bag.Own(target.Range["V10"]),publicOutput=bag.Own(target.Range["R10"]);
                            Check(Convert.ToString(secondLink.Formula).Contains("[第三家公司.xlsx]BS'!A1+K1"),"existing external address frozen while local relative reference shifts");
                            Check(Convert.ToString(publicOutput.Formula)==publicBefore,"public external formula remains unchanged");
                            Check(Convert.ToDouble(((dynamic)bag.Own(target.Range["S10"])).Value2)==123&&Convert.ToString(((dynamic)bag.Own(target.Range["T10"])).Value2)=="说明! 请检查","plain numbers and text copied unchanged");
                            dynamic literalOutput=bag.Own(target.Range["U10"]);Check(!Convert.ToBoolean(literalOutput.HasFormula)&&Convert.ToString(literalOutput.Value2)=="=1+1","literal equals text never becomes an Excel formula");
                            Call(host,"Undo");Check(((dynamic)bag.Own(target.Range["Q10"])).Value2==null&&Convert.ToString(compiledSource.Formula)==compiledBefore,"compiled Area undo restores targets and leaves source unchanged");
                            publicSource.Formula=publicBefore+"+1";reviewBlocked=false;try{Call(host,"Preview");}catch(InvalidOperationException){reviewBlocked=true;}Check(reviewBlocked,"changed source formula requires renewed review");
                            publicSource.Value2="公共说明";view.ReferenceSources.Rows.Clear();Call(host,"Preview");Check((bool)Get(host,"ready")&&view.ReferenceSources.Rows.Count==1&&Convert.ToString(view.ReferenceSources.Rows[0].Cells["decision"].Value)=="参与公司替换","single Lead-matched source is automatically identified");
                            // Linked O/P company source + vertical N years; independent of any elimination sheet.
                            view.HeadersEnabled.Checked=true;view.FormulaHeaders.Checked=true;view.RepeatNames.Checked=true;view.ReplaceHeaderArrays.Checked=false;
                            dynamic newLead=bag.Own(lead.Range["O30:P32"]);newLead.Value2=new object[,]{{"甲简称","！01-示例（集团）"},{"乙简称","！02-示例乙"},{"丙简称","第三家公司"}};
                            dynamic verticalYears=bag.Own(lead.Range["N30:N34"]);verticalYears.Value2=new object[,]{{"2027年度"},{"2026年度"},{"2025年度"},{"2024年度"},{"2023年度"}};
                            ((List<SummaryAreaPlan>)Get(host,"leadAreas")).Clear();Set(host,"lead",new SummaryAreaPlan{Sheet="Lead",Row=30,Column=15,Rows=3,Columns=2});view.Companies.Rows.Clear();
                            Set(host,"yearSource",new SummaryAreaPlan{Sheet="Lead",Row=30,Column=14,Rows=5,Columns=1});
                            // Existing full-row CSE has five-year repetitions, matching LET/SEQUENCE output semantics.
                            dynamic oldNames=bag.Own(target.Range["Q9:AE9"]);oldNames.FormulaArray="=INDEX(Lead!$O$30:$O$32,INT((COLUMN(Q9:AE9)-COLUMN(Q9))/5)+1)";oldNames.Calculate();string oldNamesFormula=Convert.ToString(((dynamic)bag.Own(target.Range["Q9"])).Formula);
                            dynamic oldYearArray=bag.Own(target.Range["Q12:U12"]);oldYearArray.FormulaArray="=TRANSPOSE(Lead!$N$30:$N$34)";oldYearArray.Calculate();string priorYearFormula=Convert.ToString(((dynamic)bag.Own(target.Range["Q12"])).Formula);
                            reviewBlocked=false;try{Call(host,"Preview");}catch(InvalidOperationException){reviewBlocked=true;}Check(reviewBlocked,"existing header arrays require explicit full-range rewrite consent");
                            view.Companies.Rows[2].Cells["selected"].Value=false;view.ReplaceHeaderArrays.Checked=true;Call(host,"Preview");Check((bool)Get(host,"ready")&&((List<SummaryHeaderArray>)Get(host,"oldHeaderArrays")).Count==2,"vertical years and existing full company/year arrays pass preflight");
                            Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);Check(!(bool)Get(host,"failed")&&(bool)Get(host,"pending"),"linked headers replace existing arrays without LET dependencies");
                            bool linked=true;for(int i=0;i<2;i++)for(int c=0;c<5;c++){dynamic name=bag.Own(target.Cells[9,17+i*5+c]);linked &= Convert.ToBoolean(name.HasFormula)&&Convert.ToString(name.Value2)==(i==0?"甲简称":"乙简称");}
                            Check(linked,"each selected company name repeats across all year columns using Lead O links");
                            dynamic newYearAnchor=bag.Own(target.Range["Q12"]);Check(Convert.ToBoolean(newYearAnchor.HasArray)&&Convert.ToString(newYearAnchor.Formula).Contains("TRANSPOSE")&&Convert.ToString(newYearAnchor.Value2)=="2027年度","vertical N years written as true TRANSPOSE array formula");
                            Check(Convert.ToString(((dynamic)bag.Own(target.Range["AA9"])).Value2)=="丙简称"&&Convert.ToBoolean(((dynamic)bag.Own(target.Range["AA9"])).HasFormula),"unselected company outside current output keeps its Lead mapping after whole-array rewrite");
                            dynamic changedCode=bag.Own(lead.Range["O30"]),changedYear=bag.Own(lead.Range["N30"]);changedCode.Value2="甲简称更新";changedYear.Value2="2030年度";target.Calculate();Check(Convert.ToString(((dynamic)bag.Own(target.Range["Q9"])).Value2)=="甲简称更新"&&Convert.ToString(newYearAnchor.Value2)=="2030年度","generated headers update when Lead source values change");changedCode.Value2="甲简称";changedYear.Value2="2027年度";target.Calculate();
                            Call(host,"Undo");Check(Convert.ToBoolean(oldNames.HasArray)&&Convert.ToString(((dynamic)bag.Own(target.Range["Q9"])).Formula)==oldNamesFormula&&Convert.ToString(newYearAnchor.Formula)==priorYearFormula,"undo restores complete original company and year array formulas");
                            var linkedPreset=host.CapturePreset();host.ApplyPreset(linkedPreset);Check(view.FormulaHeaders.Checked&&view.RepeatNames.Checked&&!view.ReplaceHeaderArrays.Checked,"formula and repeat settings persist but replacement permission resets");
                            object priorYearFormat=newYearAnchor.NumberFormat;changedYear.Value2=46388;changedYear.NumberFormat="yyyy\"年\"";string expectedDateDisplay=Convert.ToString(changedYear.Text);view.ReplaceHeaderArrays.Checked=true;
                            Call(host,"Preview");Call(host,"Start");int dateSteps=((List<SummaryAreaPlan>)Get(host,"targets")).Count;for(int i=0;i<dateSteps;i++)Call(host,"Tick",null,EventArgs.Empty);
                            Check(!(bool)Get(host,"failed")&&Convert.ToBoolean(newYearAnchor.HasFormula)&&Convert.ToDouble(newYearAnchor.Value2)==46388&&Convert.ToString(newYearAnchor.Text)==expectedDateDisplay,"linked TRANSPOSE retains numeric dates with source display format");
                            Call(host,"Undo");Check(Convert.ToString(newYearAnchor.NumberFormat)==Convert.ToString(priorYearFormat)&&Convert.ToString(newYearAnchor.Formula)==priorYearFormula,"date header undo restores original format and array formula");changedYear.Value2="2027年度";changedYear.NumberFormat="General";






                        }
                    }
                }
                if(args.Length==2&&args[0]=="--provided-header-vectors")ProvidedHeaders((object)app,session,args[1]);
                FixedHeaders((object)app,session);
                Console.WriteLine("PASS: "+checks+" Excel self-owned sample checks; no user workbook or real button interaction");return session.Finish(0);
            }
            catch(Exception e){Console.WriteLine(e);return session.Finish(1);}
        }
    }
}

[System.Runtime.Serialization.DataContract] public sealed class ProvidedHeaderVector
{
    [System.Runtime.Serialization.DataMember] public string sheet;
    [System.Runtime.Serialization.DataMember] public List<ProvidedHeaderYear> years;
    [System.Runtime.Serialization.DataMember] public List<ProvidedHeaderCompany> companies;
}
[System.Runtime.Serialization.DataContract] public sealed class ProvidedHeaderYear
{
    [System.Runtime.Serialization.DataMember] public int row;
    [System.Runtime.Serialization.DataMember] public string text;
}
[System.Runtime.Serialization.DataContract] public sealed class ProvidedHeaderCompany
{
    [System.Runtime.Serialization.DataMember] public int row;
    [System.Runtime.Serialization.DataMember] public string code;
    [System.Runtime.Serialization.DataMember] public string name;
}
