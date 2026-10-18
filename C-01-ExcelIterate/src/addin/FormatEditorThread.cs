using System;
using System.Threading;
using System.Windows.Forms;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    internal static class FormatEditorThread
    {
        private static readonly object sync=new object();
        private static Thread thread;private static DialogFormatEditor editor;private static bool closing;
        internal static bool HiddenForProbe;
        internal static Action CreatedForProbe;
        internal static IntPtr EditorHandle {get{lock(sync)return editor==null||!editor.IsHandleCreated?IntPtr.Zero:editor.Handle;}}
        internal static bool HasUnsavedDraft {get {lock(sync)return editor!=null&&editor.HasUnsavedDraft;}}
        internal static int ThreadId {get {lock(sync)return thread==null?0:thread.ManagedThreadId;}}
        private sealed class Owner:IWin32Window {private readonly IntPtr handle;internal Owner(IntPtr value){handle=value;}public IntPtr Handle {get{return handle;}}}
        internal static void Open(IntPtr owner)
        {
            lock(sync)
            {
                if(thread!=null){var active=editor;if(active!=null&&active.IsHandleCreated&&!closing)try{active.BeginInvoke((MethodInvoker)delegate{if(!active.IsDisposed&&!HiddenForProbe)active.Activate();});}catch(InvalidOperationException){}return;}
                closing=false;thread=new Thread(delegate()
                {
                    DialogFormatEditor window=null;Panel parking=null;
                    try
                    {
                        window=new DialogFormatEditor(false);
                        // Close may arrive after registration but before a native handle
                        // exists. Preserve it until Load can post to this STA loop.
                        window.Load+=delegate {lock(sync){if(closing)QueueClose(window);}};
                        lock(sync){editor=window;if(closing)return;}
                        if(CreatedForProbe!=null)CreatedForProbe();
                        if(HiddenForProbe){parking=new Panel {Size=window.Size};window.TopLevel=false;window.FormBorderStyle=FormBorderStyle.None;parking.Controls.Add(window);window.Show();}
                        using(var context=new ApplicationContext(window))
                        {if(!HiddenForProbe&&owner!=IntPtr.Zero)window.Show(new Owner(owner));Application.Run(context);}
                    }
                    catch(Exception error){Connect.Log("格式编辑线程失败："+error.GetType().Name+" HRESULT="+error.HResult.ToString("X8"));}
                    finally{if(window!=null)window.Dispose();if(parking!=null)parking.Dispose();lock(sync){editor=null;thread=null;closing=false;}}
                });thread.IsBackground=true;thread.Name="Addin format editor";thread.SetApartmentState(ApartmentState.STA);thread.Start();
            }
        }
        private static void QueueClose(DialogFormatEditor active)
        {if(active!=null&&active.IsHandleCreated&&!active.IsDisposed)try{active.BeginInvoke((MethodInvoker)delegate{if(!active.IsDisposed)active.Close();});}catch(InvalidOperationException){}}
        internal static void Close()
        {lock(sync){closing=true;QueueClose(editor);}}
    }
}
