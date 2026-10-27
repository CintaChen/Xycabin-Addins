using System;
using System.IO;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
#if WPS_WORD
using WordUpdater;
using F = WordUpdater.Tools.UpdateProbe;
using L = WordUpdater.Tools.LinkViewProbe;
#else
using Excel_Iterate;
#endif

public static class RuntimeQualityProbe
{
    private static int checks;
    private const BindingFlags All=BindingFlags.Instance|BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic;
    private static void Check(bool value,string name) {if(!value)throw new InvalidOperationException(name);checks++;Console.WriteLine("PASS "+name);}
#if WPS_WORD
    public sealed class FaultApp : F.FakeApplication
    {
        public string Fault;
        public int AlertWrites,ScreenWrites;
        public object Alerts=2,Screen=true;
        public new object DisplayAlerts {
            get {if(Fault=="alerts-read")throw new Exception("alerts read");return Alerts;}
            set {AlertWrites++;Alerts=value;if(Fault=="alerts-set"&&AlertWrites==1||Fault=="alerts-restore"&&AlertWrites>1)throw new Exception("alerts write");}
        }
        public new object ScreenUpdating {
            get {if(Fault=="screen-read")throw new Exception("screen read");return Screen;}
            set {ScreenWrites++;Screen=value;if(Fault=="screen-set"&&ScreenWrites==1||Fault=="screen-restore"&&ScreenWrites>1)throw new Exception("screen write");}
        }
    }
    private static void WordStates()
    {
        foreach(string fault in new[]{"alerts-read","screen-read","alerts-set","screen-set","alerts-restore","screen-restore","none"})
        {
            var app=new FaultApp {Fault=fault};var field=new F.FakeField("10","10",10);app.Selection.Range.Fields.Add(field);
            var result=UpdateEngine.Execute(app);
            if(fault.EndsWith("-read"))Check(app.AlertWrites==0&&app.ScreenWrites==0&&!field.UpdateCalled,"capture failure performs zero mutations: "+fault);
            else Check(Equals(app.Alerts,2)&&Equals(app.Screen,true),"every touched setting restored independently: "+fault);
            Check(result.Status==(fault=="none"?RunStatus.Completed:RunStatus.Error),"state failures are reported: "+fault);
            if(fault.EndsWith("-set"))Check(!field.UpdateCalled,"partial suppression failure prevents link update: "+fault);
        }
        var document=new L.FakeDocument();var link=new L.FakeField("LINK Excel.Sheet.12 \"C:\\sample.xlsx\" \"Sheet1!R1C1\"",10);document.Main.Fields.Add(link);
        var view=new FaultViewApp(document);var update=LinkViewEngine.UpdateAll(view);
        Check(!update.Executed&&!link.UpdateCalled&&view.Writes==0,"link viewer also stops before any mutations when state capture fails");
    }
    public sealed class FaultViewApp : L.FakeApplication
    {
        public int Writes;
        public FaultViewApp(L.FakeDocument document):base(document){}
        public new object ScreenUpdating {get{throw new Exception("screen read");}set{Writes++;}}
        public new object DisplayAlerts {get{return 2;}set{Writes++;}}
    }
#endif
    private static void Dialogs()
    {
        DialogSizing.Initialize();
        using(var grid=new ResultGrid())Check((bool)typeof(Control).GetMethod("GetStyle",All).Invoke(grid,new object[]{ControlStyles.OptimizedDoubleBuffer}),"result surface is buffered");
#if !WPS_WORD
        foreach(EliminationKind kind in new[]{EliminationKind.Receivable,EliminationKind.RevenueCost})
        using(var form=new EliminationForm(kind))
        {
            int requests=0;form.SheetsRequested+=delegate(string path){requests++;form.SetSheets(new System.Collections.Generic.List<string>{"Sheet1"},null);};
            form.SetSource("offline.xlsx","Sheet1");Check(requests==1,"source subscription populates once before load");
            typeof(EliminationForm).GetMethod("OnLoad",All).Invoke(form,new object[]{EventArgs.Empty});
            Check(requests==1,"load does not duplicate successful source population");
            Check(form.WindowState==FormWindowState.Maximized&&!form.Visible,"maximized state is established before visibility");
            var bounds=form.Bounds;typeof(EliminationForm).GetMethod("OnShown",All).Invoke(form,new object[]{EventArgs.Empty});
            Check(form.Bounds==bounds,"Shown never resets window geometry");
            form.SetPreviewMatrix(new EliminationMatrix());
            Check(!form.Execute.Enabled,"empty preview never enables writing");
        }
        using(var preview=DialogPreviewFactory.Create("elimination.receivable",new Size(1180,700)))
        {
            typeof(EliminationForm).GetMethod("OnLoad",All).Invoke(preview,new object[]{EventArgs.Empty});
            Check(preview.WindowState==FormWindowState.Normal,"size preview stays normal while production starts maximized");
        }
#endif
    }
    [STAThread] private static int Main()
    {
        string scratch=Path.Combine(Path.GetTempPath(),"runtime-quality-"+Guid.NewGuid().ToString("N"));
        DialogSizeStore.RootOverride=scratch;FormatSession.SourceOverride=Path.Combine(scratch,"missing.xml");UiDiagnostics.RootOverride=scratch;
        try
        {
            Application.EnableVisualStyles();
#if WPS_WORD
            WordStates();
#else
            checks+=ScanQualityChecks.Run();
#endif
            Dialogs();
            Check(Application.OpenForms.Count==0,"fixtures never display windows or leave open forms");
            Console.WriteLine("RuntimeQualityProbe: "+checks+" PASS, 0 FAIL");return 0;
        }
        catch(Exception error){Console.WriteLine("FAILED: "+error);return 1;}
    }
}
