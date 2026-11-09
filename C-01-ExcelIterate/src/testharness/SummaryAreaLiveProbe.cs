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
    [STAThread] private static int Main()
    {
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
                    foreach(string name in new[]{"！01-示例（集团）","！02-示例乙"})
                    {
                        dynamic source=books.Add(); bool closed=false;
                        session.Own((object)source,delegate{if(!closed)source.Close(false);});
                        using(var refs=new ComBag()) { dynamic sheets=refs.Own(source.Worksheets), sheet=refs.Own(sheets.Item(1));sheet.Name="BS";dynamic cell=refs.Own(sheet.Range["A1"]);cell.Value2=25; }
                        source.SaveAs(session.Output(name+".xlsx"),51);source.Close(false);closed=true;
                    }
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
                        var view=(SummaryAreaView)form.Controls[0];view.Folder.Text=session.DirectoryPath;
                        using(var host=new SummaryAreaHost((object)app,form,view,null,delegate{return true;}))
                        {
                            Set(host,"book",(object)book);Set(host,"bookName",Convert.ToString(book.FullName));Set(host,"templateSheet","汇总");Set(host,"pasteSheet","汇总");Set(host,"leadSheet","Lead");
                            var areas=(List<SummaryAreaPlan>)Get(host,"areas");areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=3,Column=12,Rows=1,Columns=2,Address="L3:M3"});areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=6,Column=12,Rows=1,Columns=2,Address="L6:M6"});
                            Set(host,"paste",new SummaryAreaPlan{Row=3,Column=17,Rows=1,Columns=1});Set(host,"lead",new SummaryAreaPlan{Sheet="Lead",Row=1,Column=1,Rows=2,Columns=2,Address="A1:B2"});
                            Call(host,"Preview");Check((bool)Get(host,"ready"),"read-only preflight succeeds");Check(Convert.ToInt32(worksheets.Count)==before,"preflight creates no backup sheet");Check(Convert.ToString(old.Formula)==oldFormula,"preflight leaves target formula untouched");
                            Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);
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
                        }
                    }
                }
                Console.WriteLine("PASS: "+checks+" Excel self-owned sample checks; no user workbook or real button interaction");return session.Finish(0);
            }
            catch(Exception e){Console.WriteLine(e);return session.Finish(1);}
        }
    }
}
