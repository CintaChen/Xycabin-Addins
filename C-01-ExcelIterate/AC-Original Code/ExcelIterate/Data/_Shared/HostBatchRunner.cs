using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace Excel_Iterate
{
    internal interface IHostBatchReader<T> : IDisposable
    {
        bool CanYield { get; }
        T Current { get; }
        bool MoveNext();
        void Validate();
    }

    // The producer and its COM references never leave the creating host thread.
    // Consumer owns only the bounded queue, pure processing delegates and token.
    internal sealed class HostBatchRunner<T,TResult> : IDisposable
    {
        private sealed class Consumer
        {
            internal BlockingCollection<T> Queue;
            internal CancellationToken Token;
            internal Action<T,CancellationToken> Process;
            internal Func<TResult> Finish;
            internal TResult Run()
            {
                try { foreach(var item in Queue.GetConsumingEnumerable(Token)){Token.ThrowIfCancellationRequested();Process(item,Token);} }
                catch(OperationCanceledException) {if(!Token.IsCancellationRequested)throw;}
                return Finish();
            }
        }
        private readonly int hostThread=Thread.CurrentThread.ManagedThreadId;
        private readonly IHostBatchReader<T> reader;
        private readonly BlockingCollection<T> queue=new BlockingCollection<T>(2);
        private readonly CancellationTokenSource cancellation=new CancellationTokenSource();
        private readonly Action<TResult,Exception,bool> completed;
        private readonly Action<T> progress;
        private readonly System.Windows.Forms.Timer timer=new System.Windows.Forms.Timer {Interval=15};
        private readonly Stopwatch clock=Stopwatch.StartNew();
        private long progressAt=-100;
        private Task<TResult> worker;
        private Exception failure;
        private bool producing=true,disposed,finished;
        internal int PeakQueued {get;private set;}
        internal Action Captured;
        internal bool IsRunning {get{return !disposed&&!finished;}}
        internal HostBatchRunner(IHostBatchReader<T> source,Action<T,CancellationToken> process,Func<TResult> result,Action<TResult,Exception,bool> done,Action<T> report)
        {
            reader=source;completed=done;progress=report;
            var consumer=new Consumer {Queue=queue,Token=cancellation.Token,Process=process,Finish=result};
            worker=Task.Factory.StartNew<TResult>(consumer.Run,CancellationToken.None,TaskCreationOptions.None,TaskScheduler.Default);
            timer.Tick+=delegate {Pump();};
        }
        private void AssertHost() {if(Thread.CurrentThread.ManagedThreadId!=hostThread)throw new InvalidOperationException("数据读取和会话清理必须在原宿主线程执行。");}
        internal void Start()
        {
            AssertHost();
            if(!reader.CanYield)
            {
                // Hosts without change notifications keep the previous synchronous
                // capture semantics, while the bounded pure consumer still runs.
                try
                {
                    while(!cancellation.IsCancellationRequested)
                    {
                        if(worker.IsCompleted){worker.GetAwaiter().GetResult();break;}
                        reader.Validate();if(!reader.MoveNext())break;
                        while(!queue.TryAdd(reader.Current,25,cancellation.Token))
                            if(worker.IsCompleted){worker.GetAwaiter().GetResult();throw new InvalidOperationException("数据解析提前结束。");}
                        PeakQueued=Math.Max(PeakQueued,queue.Count);Report(reader.Current);
                    }
                    EndProduction();
                }
                catch(Exception error){Fail(error);}
            }
            timer.Start();
        }
        private void Report(T block) {if(progress!=null&&clock.ElapsedMilliseconds-progressAt>=100){progressAt=clock.ElapsedMilliseconds;progress(block);}}
        private void EndProduction() {if(!producing)return;producing=false;queue.CompleteAdding();if(!cancellation.IsCancellationRequested&&failure==null&&Captured!=null)Captured();}
        private void Fail(Exception error) {if(failure==null)failure=error;cancellation.Cancel();EndProduction();}
        internal void Pump()
        {
            AssertHost();if(disposed||finished)return;
            try
            {
                if(worker.IsCompleted){Finish();return;}
                if(cancellation.IsCancellationRequested){EndProduction();return;}
                reader.Validate();
                if(producing&&queue.Count<2)
                {
                    if(reader.MoveNext()){queue.Add(reader.Current);PeakQueued=Math.Max(PeakQueued,queue.Count);Report(reader.Current);}
                    else EndProduction();
                }
            }
            catch(Exception error){Fail(error);}
        }
        private void Finish()
        {
            TResult result=default(TResult);
            try {result=worker.GetAwaiter().GetResult();reader.Validate();}
            catch(Exception error){if(failure==null&&!(error is OperationCanceledException&&cancellation.IsCancellationRequested))failure=error;}
            bool cancelled=cancellation.IsCancellationRequested&&failure==null;
            finished=true;timer.Stop();
            try {reader.Dispose();}catch(Exception error){if(failure==null)failure=error;}
            timer.Dispose();queue.Dispose();cancellation.Dispose();disposed=true;
            completed(result,failure,cancelled);
        }
        internal void Cancel() {AssertHost();if(!disposed&&!finished){cancellation.Cancel();EndProduction();}}
        public void Dispose()
        {
            AssertHost();if(disposed)return;disposed=true;timer.Stop();timer.Dispose();cancellation.Cancel();EndProduction();
            try {reader.Dispose();}
            finally
            {
                // Cleanup continuation owns managed objects only, never the reader.
                var blocks=queue;var token=cancellation;
                worker.ContinueWith(delegate(Task<TResult> task){var observed=task.Exception;blocks.Dispose();token.Dispose();},TaskScheduler.Default);
            }
        }
    }
}
