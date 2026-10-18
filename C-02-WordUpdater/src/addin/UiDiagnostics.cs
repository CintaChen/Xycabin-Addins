using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;
#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    internal static class UiDiagnostics
    {
        private static readonly Queue<string> memory=new Queue<string>();private static readonly object sync=new object();
        internal static string RootOverride;private static string digest;
        internal static string Text {get{lock(sync)return string.Join(Environment.NewLine,memory.ToArray());}}
        private static string Identity(){
            if(digest==null)try{using(var file=File.OpenRead(typeof(Connect).Assembly.Location))using(var sha=SHA256.Create())digest=BitConverter.ToString(sha.ComputeHash(file)).Replace("-","");}catch{digest="unavailable";}
            return " pid="+Process.GetCurrentProcess().Id+" thread="+Thread.CurrentThread.ManagedThreadId+" assembly="+digest;
        }
        internal static void RecordState(string stage,uint nativeThread,IntPtr owner,bool enabled,IntPtr active,IntPtr focus,IntPtr capture,uint flags){lock(sync){
            Write(DateTime.UtcNow.ToString("o")+" stage="+stage+Identity()+" native-thread="+nativeThread+" owner="+owner.ToInt64().ToString("X")+" owner-enabled="+enabled+" active="+active.ToInt64().ToString("X")+" focus="+focus.ToInt64().ToString("X")+" capture="+capture.ToInt64().ToString("X")+" gui-flags="+flags.ToString("X"));
        }}
        internal static void Record(string stage,Exception error){lock(sync){
            // Stack frames only: Message, Data, requests and recognized text are excluded.
            string row=DateTime.UtcNow.ToString("o")+" stage="+stage+" type="+error.GetType().Name+" HRESULT="+error.HResult.ToString("X8")+Identity()+Environment.NewLine+(error.StackTrace??"-");
            Write(row);
        }}
        private static void Write(string row){
            memory.Enqueue(row);while(memory.Count>20)memory.Dequeue();
            try{string path=Path.Combine(RootOverride??Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),typeof(Connect).Assembly.GetName().Name,"UiDiagnostics.log");Directory.CreateDirectory(Path.GetDirectoryName(path));if(File.Exists(path)&&new FileInfo(path).Length>256*1024)File.WriteAllText(path,"",System.Text.Encoding.UTF8);File.AppendAllText(path,row+Environment.NewLine,System.Text.Encoding.UTF8);}
            catch(Exception failed){memory.Enqueue("diagnostic-write type="+failed.GetType().Name+" HRESULT="+failed.HResult.ToString("X8"));while(memory.Count>20)memory.Dequeue();}
        }
    }
    internal sealed class ResultGrid:DataGridView,IGridColumnFit
    {
        internal Action ColumnFitOverride;
        bool IGridColumnFit.TryFitColumns() { if (ColumnFitOverride == null) return false; ColumnFitOverride(); return true; }
        internal Action<Exception> PaintFailure;internal bool PaintFailed;
        internal Action PaintForProbe;
        protected override void OnPaint(PaintEventArgs e){
            if(!PaintFailed)try{if(PaintForProbe!=null)PaintForProbe();base.OnPaint(e);return;}catch(Exception error){PaintFailed=true;UiDiagnostics.Record("file.names-paint",error);if(PaintFailure!=null)PaintFailure(error);}
            e.Graphics.Clear(Color.White);TextRenderer.DrawText(e.Graphics,"结果显示失败，识别结果已保留。请复制诊断，并重新打开窗口重试。",SystemFonts.MessageBoxFont,ClientRectangle,Color.Firebrick,TextFormatFlags.WordBreak|TextFormatFlags.NoPrefix);
        }
    }
}
