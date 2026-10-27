using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using Excel_Iterate;

public static class ScanQualityChecks
{
    private static int checks;
    private static void Check(bool value,string text){if(!value)throw new Exception(text);checks++;Console.WriteLine("PASS "+text);}
    public sealed class Book {public string Name="Book.xlsx",FullName="Book.xlsx";}
    public sealed class Books
    {
        public readonly List<object> Items=new List<object>();
        public int Count {get{return Items.Count;}}
        public object Item(int index){return Items[index-1];}
    }
    public sealed class ApplicationBooks {public readonly Books Workbooks=new Books();}
    public sealed class Sheet {public object Parent;}
    public sealed class Point : IConnectionPoint
    {
        public IDirectoryAppEvents Sink;public bool Released;
        public void Advise(object value,out int cookie){Sink=(IDirectoryAppEvents)value;cookie=1;}
        public void Unadvise(int cookie){Released=true;Sink=null;}
        public void GetConnectionInterface(out Guid id){id=typeof(IDirectoryAppEvents).GUID;}
        public void GetConnectionPointContainer(out IConnectionPointContainer container){container=null;}
        public void EnumConnections(out IEnumConnections connections){connections=null;}
    }
    public sealed class Events : IConnectionPointContainer
    {
        public readonly Point Point=new Point();
        public void FindConnectionPoint(ref Guid id,out IConnectionPoint point){point=Point;}
        public void EnumConnectionPoints(out IEnumConnectionPoints points){points=null;}
    }
    private sealed class Reader : IHostBatchReader<FormulaBlock>
    {
        internal int ReadThread=Thread.CurrentThread.ManagedThreadId,Count,Limit=25;
        internal bool Changed,Disposed,Yield=true;
        internal Exception ReadError;
        public bool CanYield {get{return Yield;}}
        public FormulaBlock Current {get;private set;}
        public void Validate(){if(ReadThread!=Thread.CurrentThread.ManagedThreadId)throw new Exception("COM reader crossed thread");if(Changed)throw new InvalidOperationException("source changed");}
        public bool MoveNext(){Validate();if(ReadError!=null)throw ReadError;if(Count++>=Limit)return false;Current=FormulaBlock.Capture(new object[,]{{"=A1",""}},1,2);Current.FirstRow=Count;Current.FirstColumn=1;Current.WorkbookName="Book.xlsx";Current.SheetName="Sheet1";return true;}
        public void Dispose(){ValidateThread();Disposed=true;}
        private void ValidateThread(){if(ReadThread!=Thread.CurrentThread.ManagedThreadId)throw new Exception("reader disposed on worker");}
    }
    private static void Pump(HostBatchRunner<FormulaBlock,DependentResult> session,Func<bool> done)
    {var limit=Stopwatch.StartNew();while(!done()&&limit.ElapsedMilliseconds<5000){session.Pump();Thread.Sleep(1);}Check(done(),"batch task finishes within deadline");}
    internal static int Run()
    {
        EliminationStates();GridMeasurements();
        object[,] array=(object[,])Array.CreateInstance(typeof(object),new[]{2,1},new[]{-2,3});array[-2,3]="=A1";array[-1,3]="=B1";
        Check(Equals(HostValueArray.At(array,1,0,2,1),"=B1"),"arbitrary two-dimensional array lower bounds preserved");
        Check(Equals(HostValueArray.At(new object[]{"=A1","=B1"},0,1,1,2),"=B1"),"one-dimensional row array accepted");
        Check(Equals(HostValueArray.At(new object[]{"=A1","=B1"},1,0,2,1),"=B1"),"one-dimensional column array accepted");
        var vector=Array.CreateInstance(typeof(object),new[]{2},new[]{-4});vector.SetValue("tail",-3);
        Check(Equals(HostValueArray.At(vector,1,0,2,1),"tail")&&Equals(HostValueArray.At("=A1",0,0,1,1),"=A1"),"scalar and nonzero lower-bound vector are mapped exactly");
        bool rejected=false;try{HostValueArray.At(new object[4],0,0,2,2);}catch(InvalidOperationException){rejected=true;}Check(rejected,"ambiguous flattened matrix rejected");
        foreach(bool yield in new[]{true,false})
        {
            var reader=new Reader {Yield=yield};var result=new DependentResult {TargetWorkbookName="Book.xlsx",TargetSheetName="Sheet1",TargetCellAddress="A1"};
            var accumulator=new DependentAccumulator(result);bool done=false;Exception failure=null;int processingThread=0;
            using(var session=new HostBatchRunner<FormulaBlock,DependentResult>(reader,delegate(FormulaBlock b,CancellationToken t){processingThread=Thread.CurrentThread.ManagedThreadId;accumulator.Accept(b,t);},accumulator.Finish,delegate(DependentResult r,Exception e,bool cancelled){done=true;failure=e;},null))
            {session.Start();Pump(session,()=>done);Check(failure==null&&result.Items.Count==25&&result.ScannedCells==50,"row-major findings preserved in yield/fallback mode");Check(session.PeakQueued<=2,"pending queue never exceeds two blocks");Check(processingThread!=reader.ReadThread&&reader.Disposed,"only pure parsing runs on worker and reader is released on host");}
        }
        foreach(bool sourceChange in new[]{false,true})
        {
            var reader=new Reader {Limit=100000};var result=new DependentResult();var accumulator=new DependentAccumulator(result);bool done=false,cancelledResult=false;Exception failure=null;
            using(var session=new HostBatchRunner<FormulaBlock,DependentResult>(reader,accumulator.Accept,accumulator.Finish,delegate(DependentResult r,Exception e,bool cancelled){done=true;failure=e;cancelledResult=cancelled;},null))
            {session.Start();session.Pump();if(sourceChange)reader.Changed=true;else session.Cancel();Pump(session,()=>done);Check(sourceChange?failure!=null:cancelledResult,"source changes abort; user cancellation stays distinguishable");Check(reader.Disposed,"cancelled reader released on host");}
        }
        foreach(bool comFailure in new[]{false,true})
        {
            var expected=comFailure?(Exception)new System.Runtime.InteropServices.COMException("RPC_E_CALL_REJECTED",unchecked((int)0x80010001)):new InvalidOperationException("invalid array shape");
            var reader=new Reader {ReadError=expected};var accumulator=new DependentAccumulator(new DependentResult());bool done=false;Exception reported=null;
            using(var session=new HostBatchRunner<FormulaBlock,DependentResult>(reader,accumulator.Accept,accumulator.Finish,delegate(DependentResult r,Exception e,bool cancelled){done=true;reported=e;},null))
            {session.Start();Pump(session,()=>done);Check(ReferenceEquals(reported,expected)&&reader.Disposed,"read failure is reported and released without claiming complete results");}
        }
        {
            var reader=new Reader();var accumulator=new DependentAccumulator(new DependentResult());int callbacks=0;
            using(var entered=new ManualResetEvent(false))using(var release=new ManualResetEvent(false))using(var exited=new ManualResetEvent(false))
            {
                var session=new HostBatchRunner<FormulaBlock,DependentResult>(reader,delegate(FormulaBlock b,CancellationToken t){entered.Set();release.WaitOne();try{accumulator.Accept(b,t);}finally{exited.Set();}},accumulator.Finish,delegate{callbacks++;},null);
                session.Start();session.Pump();Check(entered.WaitOne(5000),"worker fixture entered pure computation");session.Dispose();release.Set();
                Check(exited.WaitOne(5000)&&callbacks==0&&reader.Disposed,"closing a session discards late completion and releases only its reader");
            }
        }
        var events=new Events();var book=new Book();
        using(var watch=new SourceChangeWatch(events))
        {Check(watch.Available,"change monitor subscribes through the connection point");watch.WatchBook(book);events.Point.Sink.SheetChange(new Sheet {Parent=book},null);bool aborted=false;try{watch.Validate();}catch(InvalidOperationException){aborted=true;}Check(aborted,"source edit invalidates pending snapshot");}
        Check(events.Point.Released,"change monitor unadvises on completion");
        var ccw=System.Runtime.InteropServices.Marshal.GetComInterfaceForObject(new DirectoryEventSink(),typeof(IDirectoryAppEvents));
        try{Check(ccw!=IntPtr.Zero,"existing COM-visible sink supports the event interface without adding public types");}
        finally{System.Runtime.InteropServices.Marshal.Release(ccw);}
        foreach(string change in new[]{"calculate","close","delete","new-sheet"})
        using(var watch=new SourceChangeWatch(events))
        {
            watch.WatchBook(book);var sheet=new Sheet {Parent=book};bool cancel=false;
            if(change=="calculate")events.Point.Sink.SheetCalculate(sheet);
            else if(change=="close")events.Point.Sink.WorkbookBeforeClose(book,ref cancel);
            else if(change=="delete")events.Point.Sink.SheetBeforeDelete(sheet);
            else events.Point.Sink.WorkbookNewSheet(book,sheet);
            bool aborted=false;try{watch.Validate();}catch(InvalidOperationException){aborted=true;}
            Check(aborted,"source notification invalidates snapshot: "+change);
        }
        var host=new ApplicationBooks();host.Workbooks.Items.Add(book);var original=new SourceBookSnapshot(book);original.Validate(host);
        host.Workbooks.Items[0]=new Book();bool replaced=false;try{original.Validate(host);}catch(InvalidOperationException){replaced=true;}
        Check(replaced,"reopening the same workbook name cannot replace the captured object");
        using(var watch=new SourceChangeWatch(new object()))Check(!watch.Available,"unsupported host selects synchronous fallback");
        return checks;
    }
    private static void EliminationStates()
    {
        var request=new EliminationRequest();request.Columns.Add(new EliminationColumn {Role=EliminationRole.Amount});
        var copy=request.Copy();request.Columns[0].Role=EliminationRole.Ignore;
        Check(copy.Columns[0].Role==EliminationRole.Amount,"worker request does not share mutable UI column tags");
        using(var session=new EliminationSession(new object()))
        {
            session.Target=new EliminationTargetHandle {SheetName="result"};session.DiffTarget=new EliminationTargetHandle {SheetName="difference"};session.Snapshot=new MergeSnapshot();session.DiffSnapshot=new MergeSnapshot();session.Phase=EliminationPhase.UndoAvailable;
            var order=new List<string>();bool failed=false;
            try {session.RestoreTargets(delegate(EliminationTargetHandle target,MergeSnapshot snapshot){order.Add(target.SheetName);if(target.SheetName=="difference")throw new Exception("restore failure");});}catch(Exception){failed=true;}
            Check(failed&&session.Target==null&&session.DiffTarget!=null&&session.DiffSnapshot!=null,"partial undo preserves the remaining recoverable target");
            session.RestoreTargets(delegate(EliminationTargetHandle target,MergeSnapshot snapshot){order.Add(target.SheetName);});
            Check(string.Join(",",order)=="result,difference,difference"&&session.DiffTarget==null&&session.Phase==EliminationPhase.Idle,"dual-table undo retries only the unfinished table");
        }
        using(var session=new EliminationSession(new object()))
        {
            session.Target=new EliminationTargetHandle {SheetName="untouched-existing"};
            session.DiffTarget=new EliminationTargetHandle {SheetName="created",IsNewSheet=true};
            int restores=0;session.RestoreTargets(delegate(EliminationTargetHandle target,MergeSnapshot snapshot){if(target.SheetName!="created"||snapshot!=null)throw new Exception("unexpected restore");restores++;});
            Check(restores==1&&session.Target==null&&session.DiffTarget==null,"partial setup removes new sheets and leaves untouched existing targets unchanged");
        }
        var cancellation=new CancellationTokenSource();cancellation.Cancel();bool stopped=false;
        try {EliminationCancellation.Run(cancellation.Token,delegate{return 1;});}catch(OperationCanceledException){stopped=true;}
        Check(stopped,"cancelled elimination does not enter calculation");cancellation.Dispose();
    }
    private static void GridMeasurements()
    {
        using(var grid=new System.Windows.Forms.DataGridView {AllowUserToAddRows=false})
        {
            grid.Columns.Add("value","Value");grid.Rows.Add("small");grid.Rows.Add(new string('W',80));Ui.AutoFitColumns(grid);
            int full=grid.Columns[0].Width;Check(full>200,"column width includes the final off-screen data row");
            int manual=grid.Columns[0].MinimumWidth+20;grid.Columns[0].Width=manual;Ui.AutoFitColumns(grid);
            Check(grid.Columns[0].Width==manual,"unchanged data skips redundant fitting and preserves manual width");
            grid.Rows[1].Cells[0].Value=new string('W',120);Ui.AutoFitColumns(grid);
            Check(grid.Columns[0].Width>full,"data changes invalidate the measurement cache");
        }
        using(var form=new EliminationForm(EliminationKind.Receivable))
        {
            var matrix=new EliminationMatrix();matrix.Years.Add("2026");matrix.Columns.AddRange(EliminationMatrixBuilder.BuildColumns(matrix.Years,1));
            for(int i=0;i<5000;i++)matrix.Rows.Add(new MatrixRow {Entity="A",Customer="B",Note=i==4999?"last-data-row":"row"});
            form.SetPreviewMatrix(matrix);
            Check(form.Entries.VirtualMode&&form.Entries.Rows.Count==5000,"large preview keeps every row through virtual value access");
            Check(Convert.ToString(form.Entries.Rows[4999].Cells[2].Value)=="last-data-row","virtual preview returns the final row value");
            form.Entries.ClearSelection();form.Entries.Rows[4999].Selected=true;
            var clipboard=form.Entries.GetClipboardContent();Check(clipboard!=null&&clipboard.GetText().Contains("last-data-row"),"native copy includes virtual row values without touching the clipboard");
            form.SetPreviewMatrix(new EliminationMatrix());
            Check(!form.Entries.VirtualMode&&form.Entries.Rows.Count==0,"next empty result discards the virtual view");
        }
    }
}
