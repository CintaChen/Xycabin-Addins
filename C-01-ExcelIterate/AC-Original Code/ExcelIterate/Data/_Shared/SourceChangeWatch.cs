using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;

namespace Excel_Iterate
{
    internal sealed class SourceChangeWatch : IDisposable
    {
        private readonly int hostThread=Thread.CurrentThread.ManagedThreadId;
        private readonly List<object> books=new List<object>();
        private IConnectionPoint point;private DirectoryEventSink sink;private int cookie;
        private volatile bool changed;
        private volatile bool disposed;
        private readonly bool migrationMode;
        private sealed class Pending {internal object Value;internal bool Workbook;internal string Kind;}
        private readonly Queue<Pending> pending=new Queue<Pending>();
        private string reason="未知事件";
        internal string Role="来源";
        internal Action<string> Trace;
        internal bool Available {get;private set;}
        internal SourceChangeWatch(object application):this(application,false){}
        internal SourceChangeWatch(object application,bool migration)
        {
            migrationMode=migration;
            try
            {
                var container=application as IConnectionPointContainer;if(container==null)return;
                Guid id=typeof(IDirectoryAppEvents).GUID;container.FindConnectionPoint(ref id,out point);
                sink=new DirectoryEventSink {SourceEvent=OnChanged};point.Advise(sink,out cookie);Available=true;
            }
            catch(Exception error){UiDiagnostics.Record("source-monitor",error);Dispose();}
        }
        internal void WatchBook(object book){books.Add(book);}
        internal void MarkChanged() {changed=true;}
        private void OnChanged(object value,bool workbook,string kind)
        {
            if(disposed)return;
            if(Trace!=null)Trace(Role+" event="+kind+" thread="+Thread.CurrentThread.ManagedThreadId+" host="+hostThread);
            // Migration snapshots formula TEXT, never calculated results. A
            // calculation notification alone cannot invalidate that snapshot.
            // Value/formula edits still arrive as SheetChange and remain fatal.
            if(migrationMode&&kind=="SheetCalculate"){if(Trace!=null)Trace(Role+" ignored=calculation-only");return;}
            if(Thread.CurrentThread.ManagedThreadId!=hostThread)
            {
                if(migrationMode){lock(pending){if(pending.Count>=1024){reason="事件队列溢出";changed=true;}else pending.Enqueue(new Pending{Value=value,Workbook=workbook,Kind=kind});}return;}
                reason=kind+"：异线程事件，尚未核对工作簿";changed=true;return;
            }
            Handle(value,workbook,kind);
        }
        private void Handle(object value,bool workbook,string kind)
        {
            object parent=null;
            try
            {
                object book=value;if(!workbook){parent=((dynamic)value).Parent;book=parent;}
                foreach(object watched in books)if(WorksheetIdentity.SameObject(watched,book)){if(!changed)reason=kind;changed=true;if(Trace!=null)Trace(Role+" matched="+kind);return;}
            }
            catch(Exception error) {reason=kind+"：事件身份核对失败 "+error.HResult.ToString("X8");changed=true;}
            finally {if(parent!=null&&Marshal.IsComObject(parent))Marshal.ReleaseComObject(parent);}
        }
        internal void ProcessPending()
        {
            if(Thread.CurrentThread.ManagedThreadId!=hostThread)throw new InvalidOperationException("来源监听只能在原宿主线程验证。");
            while(true){Pending item;lock(pending){if(pending.Count==0)break;item=pending.Dequeue();}Handle(item.Value,item.Workbook,item.Kind);}
        }
        internal void Validate()
        {
            ProcessPending();
            if(changed)throw new InvalidOperationException(Role+"变化："+reason+"；本次任务已中止，请重新预检。");
        }
        public void Dispose()
        {
            disposed=true;lock(pending)pending.Clear();
            if(sink!=null){sink.SourceChanged=null;sink.SourceEvent=null;}
            var connection=point;point=null;
            try {if(connection!=null&&cookie!=0)connection.Unadvise(cookie);}
            finally {cookie=0;sink=null;Available=false;if(connection!=null&&Marshal.IsComObject(connection))Marshal.ReleaseComObject(connection);}
        }
    }
}
