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
        internal bool Available {get;private set;}
        internal SourceChangeWatch(object application)
        {
            try
            {
                var container=application as IConnectionPointContainer;if(container==null)return;
                Guid id=typeof(IDirectoryAppEvents).GUID;container.FindConnectionPoint(ref id,out point);
                sink=new DirectoryEventSink {SourceChanged=OnChanged};point.Advise(sink,out cookie);Available=true;
            }
            catch(Exception error){UiDiagnostics.Record("source-monitor",error);Dispose();}
        }
        internal void WatchBook(object book){books.Add(book);}
        internal void MarkChanged() {changed=true;}
        private void OnChanged(object value,bool workbook)
        {
            if(Thread.CurrentThread.ManagedThreadId!=hostThread){changed=true;return;}
            object parent=null;
            try
            {
                object book=value;if(!workbook){parent=((dynamic)value).Parent;book=parent;}
                foreach(object watched in books)if(WorksheetIdentity.SameObject(watched,book)){changed=true;return;}
            }
            catch {changed=true;}
            finally {if(parent!=null&&Marshal.IsComObject(parent))Marshal.ReleaseComObject(parent);}
        }
        internal void Validate()
        {
            if(Thread.CurrentThread.ManagedThreadId!=hostThread)throw new InvalidOperationException("来源监听只能在原宿主线程验证。");
            if(changed)throw new InvalidOperationException("来源数据已修改、重算或关闭，本次任务已中止，请重试。");
        }
        public void Dispose()
        {
            if(sink!=null)sink.SourceChanged=null;
            var connection=point;point=null;
            try {if(connection!=null&&cookie!=0)connection.Unadvise(cookie);}
            finally {cookie=0;sink=null;Available=false;if(connection!=null&&Marshal.IsComObject(connection))Marshal.ReleaseComObject(connection);}
        }
    }
}
