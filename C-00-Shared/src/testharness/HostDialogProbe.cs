using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
#if WPS_WORD
using WordUpdater;
#else
using Excel_Iterate;
#endif

internal static class HostDialogProbe
{
    [DllImport("user32.dll")] private static extern bool EnableWindow(IntPtr handle,bool enabled);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr handle);
    [DllImport("user32.dll")] private static extern IntPtr GetFocus();
    private static int checks;
    private static void Check(bool value,string name) { if(!value)throw new InvalidOperationException(name);checks++;Console.WriteLine("PASS "+name); }
    [STAThread] public static int Main()
    {
        Console.OutputEncoding=System.Text.Encoding.UTF8;
        string sandbox=Path.Combine(Path.GetTempPath(),"HostDialogProbe-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(sandbox);UiDiagnostics.RootOverride=sandbox;
        int main=Thread.CurrentThread.ManagedThreadId;IntPtr focus=GetFocus();bool success=false;
        try
        {
            HostDialogService.Register();
            using(var parking=new Panel())using(var owner=new Form {TopLevel=false})using(var dispatcher=new Control())
            {
                parking.Controls.Add(owner);var parent=parking.Handle;var handle=owner.Handle;var dispatch=dispatcher.Handle;
                SynchronizationContext.SetSynchronizationContext(null);
                int invoked=0,worker=0;bool disabled=false;int nested=0;IntPtr owned=IntPtr.Zero;
                int value=HostDialogService.Run<int>(owner,uiOwner=>{
                    worker=Thread.CurrentThread.ManagedThreadId;
                    Check(worker!=main&&Thread.CurrentThread.GetApartmentState()==ApartmentState.STA,"confirmation uses independent STA");
                    Check(!ReferenceEquals(uiOwner,owner),"UI receives HWND wrapper rather than host Form");
                    owned=uiOwner.Handle;disabled=!IsWindowEnabled(owned);
                    dispatcher.BeginInvoke((MethodInvoker)delegate {invoked++;nested=HostDialogService.Run<int>(null,unused=>99,-1);});
                    var until=DateTime.UtcNow.AddSeconds(2);while(Volatile.Read(ref invoked)==0&&DateTime.UtcNow<until)Thread.Sleep(2);
                    return 42;
                },-1);
                Check(value==42&&worker!=main,"synchronous business result returns to original STA");
                Check(invoked==1&&nested==-1,"native COM wait pumps host messages and blocks nested confirmations");
                Check(disabled&&IsWindowEnabled(owned),"only owned HWND is disabled and restored");
                Check(SynchronizationContext.Current==null&&GetFocus()==focus,"host context and desktop focus are preserved");
                EnableWindow(owned,false);
                HostDialogService.Run<int>(owner,unused=>1,0);
                Check(!IsWindowEnabled(owned),"pre-disabled owner remains disabled");EnableWindow(owned,true);
                value=HostDialogService.Run<int>(owner,unused=>{throw new InvalidOperationException("private input");},-1);
                Check(value==-1&&IsWindowEnabled(owned),"UI failure cancels safely and restores owner");
                Check(!UiDiagnostics.Text.Contains("private input"),"shared dialog diagnostics exclude exception content");
                int loops=0;
                HostDialogService.Run<int>(owner,unused=>{
                    using(var hidden=new Form {TopLevel=false})using(var container=new Panel())using(var timer=new System.Windows.Forms.Timer {Interval=15})
                    {container.Controls.Add(hidden);using(var context=new ApplicationContext(hidden)) {hidden.Show();timer.Tick+=delegate{loops++;hidden.Close();};timer.Start();Application.Run(context);}}
                    return 1;
                },0);
                Check(loops==1&&IsWindowEnabled(owned),"real hidden UI loop exits without leaving owner disabled");
                using(var entered=new ManualResetEvent(false))
                {
                    HostDialogService.WaitForProbe=delegate {entered.WaitOne(2000);return unchecked((int)0x80004005);};
                    try
                    {
                        value=HostDialogService.Run<int>(owner,unused=>{
                            using(var hidden=new Form {TopLevel=false})using(var container=new Panel())
                            {container.Controls.Add(hidden);using(var context=new ApplicationContext(hidden)){hidden.Show();entered.Set();Application.Run(context);}}
                            return 7;
                        },-1);
                    }
                    finally {HostDialogService.WaitForProbe=null;}
                    var deadline=DateTime.UtcNow.AddSeconds(5);while(HostDialogService.ActiveWorkers!=0&&DateTime.UtcNow<deadline){Application.DoEvents();Thread.Sleep(5);}
                    Check(value==-1&&HostDialogService.ActiveWorkers==0&&IsWindowEnabled(owned),"native wait failure cancels worker UI and releases owner");
                }
            }
            success=true;Console.WriteLine("HostDialogProbe: "+checks+" PASS");return 0;
        }
        catch(Exception error){Console.WriteLine("FAIL "+error);return 1;}
        finally{if(success&&Path.GetDirectoryName(Path.GetFullPath(sandbox))==Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar))Directory.Delete(sandbox,true);}
    }
}
