using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using Excel_Iterate;

internal static class SummaryAreaChecks
{
    internal static void Run(Action<string, bool, string> check)
    {
        Action<Action, string> rejects = delegate(Action action, string name) { bool rejected = false; try { action(); } catch (InvalidOperationException) { rejected = true; } check(name, rejected, "invalid input accepted"); };
        var areas = new List<SummaryAreaPlan> {
            new SummaryAreaPlan { Sheet="汇总", Row=3,Column=12,Rows=4,Columns=5 },
            new SummaryAreaPlan { Sheet="汇总", Row=20,Column=12,Rows=11,Columns=5 } };
        var targets = SummaryAreaEngine.Targets(areas, 3, 17, 2);
        check("Ctrl 第一家公司 Q3:U6", SummaryAreaEngine.Address(targets[0]) == "Q3:U6", SummaryAreaEngine.Address(targets[0]));
        check("Ctrl 竖向空隙保持 Q20:U30", SummaryAreaEngine.Address(targets[1]) == "Q20:U30", SummaryAreaEngine.Address(targets[1]));
        check("Ctrl 第二家公司 V3:Z6", SummaryAreaEngine.Address(targets[2]) == "V3:Z6", SummaryAreaEngine.Address(targets[2]));
        check("Ctrl 第二片 V20:Z30", SummaryAreaEngine.Address(targets[3]) == "V20:Z30", SummaryAreaEngine.Address(targets[3]));
        check("计划从未包含间隔行", targets.TrueForAll(a => a.Row == 3 || a.Row == 20), "gap copied");
        targets = SummaryAreaEngine.Targets(areas, 50, 20, 1);
        check("行列均平移且保留17行偏移", targets[1].Row == 67 && targets[1].Column == 20, "offset wrong");
        rejects(delegate { SummaryAreaEngine.Targets(areas,3,12,1); }, "模板重叠拒绝");
        rejects(delegate { SummaryAreaEngine.Targets(areas,3,16383,1); }, "目标列越界拒绝");
        rejects(delegate { SummaryAreaEngine.Targets(areas,1048570,17,1); }, "目标行越界拒绝");
        rejects(delegate { SummaryAreaEngine.Targets(areas,3,17,0); }, "未选公司拒绝");
        areas.Add(new SummaryAreaPlan { Sheet="其他",Row=40,Column=12,Rows=1,Columns=1 });
        rejects(delegate { SummaryAreaEngine.Targets(areas,3,17,1); }, "跨表模板拒绝"); areas.RemoveAt(2);
        areas.Add(new SummaryAreaPlan { Sheet="汇总",Row=4,Column=13,Rows=1,Columns=1 });
        rejects(delegate { SummaryAreaEngine.Targets(areas,3,17,1); }, "重叠 Area 拒绝"); areas.RemoveAt(2);
        string formula = SummaryTextReference.External(@"N:\EASDATA\项目'一\", "！01-示例（集团）", "BS!$L$17");
        check("文件名称全角字符及固定源地址保留", formula == "='N:\\EASDATA\\项目''一\\[！01-示例（集团）.xlsx]BS'!$L$17", formula);
        formula = SummaryTextReference.External(@"C:\项目", "公司甲", "ADJ'!$G$1");
        check("单引号工作表转义", formula == "='C:\\项目\\[公司甲.xlsx]ADJ'''!$G$1", formula);
        formula = SummaryTextReference.External(@"C:\项目", "公司甲", "'工作 表''一'!G1");
        check("带引号空格中文工作表", formula.EndsWith("]工作 表''一'!G1", StringComparison.Ordinal), formula);
        foreach (string invalid in new[] { "标题文字", "BS!XFE1", "BS!A1048577", "BS!A0", "BS!A1:B2", "=SUM(A1)", "BS!A1+1" })
        { string text=invalid; rejects(delegate { SummaryTextReference.External(@"C:\项目", "甲", text); }, "无效引用阻断 "+text); }
        rejects(delegate { SummaryTextReference.External(@"C:\项目", "甲/乙", "BS!A1"); }, "非法公司名拒绝而不裁剪");
        rejects(delegate { SummaryTextReference.External("relative", "甲", "BS!A1"); }, "相对文件夹拒绝");
        SummaryAreaEngine.ValidateShift("=ROUND(F76-F121,2)", 20, 5);
        check("ROUND 原生公式允许2D平移", true, "");
        rejects(delegate { SummaryAreaEngine.NativeFormula("=INDIRECT(A1)"); }, "INDIRECT 阻断");
        rejects(delegate { SummaryAreaEngine.NativeFormula("='C:\\[a.xlsx]BS'!A1"); }, "不明确已有外链阻断");
        rejects(delegate { SummaryAreaEngine.NativeFormula("=CustomName+1"); }, "未确认名称阻断");
        rejects(delegate { SummaryAreaEngine.ValidateShift("=XFD1",0,1); }, "原生相对引用越界预检阻断");
        using (var form = new SingleSummaryForm(true))
        {
            check("Ribbon MVP使用四组参数视图", form.Controls.Count==1 && form.Controls[0] is SummaryAreaView, "view missing");
            var view=(SummaryAreaView)form.Controls[0]; check("来源文件检查默认关闭",!view.CheckFiles.Checked,"default checked");
            var app=new SummaryAreaFakeApp();
            using(var host=new SummaryAreaHost(app,form,view))
            {
                var flags=BindingFlags.Instance|BindingFlags.NonPublic;
                Action<string,object> field=(name,value)=>typeof(SummaryAreaHost).GetField(name,flags).SetValue(host,value);
                field("book",app.Book);field("bookName","test.xlsx");field("templateSheet","汇总");field("backup","备份");field("targets",SummaryAreaEngine.Targets(areas,3,17,2));
                var restore=typeof(SummaryAreaHost).GetMethod("Restore",flags);
                restore.Invoke(host,new object[]{4});
                check("生产回滚只访问四个Area",app.Book.Target.Copied.Count==4 && app.Book.Target.Copied[0]=="Q3:U6" && app.Book.Target.Copied[1]=="Q20:U30" && app.Book.Target.Copied[3]=="V20:Z30","rollback touched gaps");
                check("回滚恢复宿主状态",app.EnableEvents&&app.ScreenUpdating&&app.AskToUpdateLinks&&app.DisplayAlerts,"host state leaked");
                app.Book.Target.Fail=true;bool failed=false;try{restore.Invoke(host,new object[]{1});}catch(TargetInvocationException){failed=true;}
                check("回滚写入失败明确抛出",failed,"restore failure swallowed");
                check("失败仍恢复宿主状态",app.EnableEvents&&app.ScreenUpdating&&app.AskToUpdateLinks&&app.DisplayAlerts,"failed host state leaked");
                field("pending",true);field("failed",true);field("step",4);
                bool blocked=false;try{typeof(SummaryAreaHost).GetMethod("Confirm",flags).Invoke(host,null);}catch(TargetInvocationException){blocked=true;}
                check("失败批次不能确认丢失恢复信息",blocked,"failed session confirmed");
            }
        }
    }
}
public sealed class SummaryAreaFakeApp
{
    public bool EnableEvents{get;set;} public bool ScreenUpdating{get;set;} public bool AskToUpdateLinks{get;set;} public bool DisplayAlerts{get;set;}
    public SummaryAreaFakeBook Book=new SummaryAreaFakeBook();
    public SummaryAreaFakeApp(){EnableEvents=ScreenUpdating=AskToUpdateLinks=DisplayAlerts=true;}
    public SummaryAreaFakeBooks Workbooks{get{return new SummaryAreaFakeBooks(Book);}}
}
public sealed class SummaryAreaFakeBooks
{
    private readonly SummaryAreaFakeBook book;public SummaryAreaFakeBooks(SummaryAreaFakeBook book){this.book=book;}public int Count{get{return 1;}}public object Item(int index){return book;}
}
public sealed class SummaryAreaFakeBook
{
    public string FullName{get{return "test.xlsx";}}public SummaryAreaFakeSheet Target=new SummaryAreaFakeSheet();public SummaryAreaFakeSheet Original=new SummaryAreaFakeSheet();public SummaryAreaFakeSheets Worksheets{get{return new SummaryAreaFakeSheets(this);}}
}
public sealed class SummaryAreaFakeSheets
{
    private readonly SummaryAreaFakeBook book;public SummaryAreaFakeSheets(SummaryAreaFakeBook b){book=b;}public object Item(string name){return name=="备份"?book.Original:book.Target;}
}
public sealed class SummaryAreaFakeSheet
{
    public readonly List<string> Copied=new List<string>();public bool Fail;public SummaryAreaFakeRanges Range{get{return new SummaryAreaFakeRanges(this);}}
}
public sealed class SummaryAreaFakeRanges
{
    private readonly SummaryAreaFakeSheet sheet;public SummaryAreaFakeRanges(SummaryAreaFakeSheet s){sheet=s;}public SummaryAreaFakeRange this[string address]{get{return new SummaryAreaFakeRange(sheet,address);}}
}
public sealed class SummaryAreaFakeRange
{
    private readonly SummaryAreaFakeSheet sheet;private readonly string address;public SummaryAreaFakeRange(SummaryAreaFakeSheet s,string a){sheet=s;address=a;}public void Clear(){if(sheet.Fail)throw new InvalidOperationException("injected restore failure");}public void Copy(object target){var r=(SummaryAreaFakeRange)target;r.sheet.Copied.Add(r.address);}
}
