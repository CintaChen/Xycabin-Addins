using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using WpsShared;

internal static class RestartHelper
{
    [DllImport("oleaut32.dll",PreserveSig=false)]private static extern void GetActiveObject(ref Guid clsid,IntPtr reserved,[MarshalAs(UnmanagedType.IUnknown)]out object value);
    [DllImport("user32.dll")]private static extern uint GetWindowThreadProcessId(IntPtr handle,out uint pid);
    private static int Pid(object app) {IntPtr handle;try{handle=new IntPtr(Convert.ToInt64(((dynamic)app).Hwnd));}catch{object window=((dynamic)app).ActiveWindow;try{handle=new IntPtr(Convert.ToInt64(((dynamic)window).Hwnd));}finally{Marshal.ReleaseComObject(window);}}uint pid;GetWindowThreadProcessId(handle,out pid);return (int)pid;}
    [STAThread]private static int Main(string[] args) {
        RestartTicket ticket=null;object app=null;
        try {
            if(args.Length!=1)return 2;ticket=RestartTicket.Load(args[0]);
            using(var source=Process.GetProcessById(ticket.ProcessId)) {
                if(source.StartTime.ToUniversalTime().Ticks!=ticket.Started||source.SessionId!=ticket.SessionId||!string.Equals(source.MainModule.FileName,ticket.Executable,StringComparison.OrdinalIgnoreCase))throw new IOException("原WPS进程身份不匹配。");
                using(var ready=EventWaitHandle.OpenExisting("Local\\Wps.Restart."+ticket.Id+".ready"))using(var go=EventWaitHandle.OpenExisting("Local\\Wps.Restart."+ticket.Id+".go"))using(var abort=EventWaitHandle.OpenExisting("Local\\Wps.Restart."+ticket.Id+".abort")) {
                    ready.Set();if(WaitHandle.WaitAny(new WaitHandle[]{go,abort},30000)!=0)return 3;
                    var until=DateTime.UtcNow.AddSeconds(60);while(!source.WaitForExit(200)){if(abort.WaitOne(0))return 3;if(DateTime.UtcNow>=until)throw new IOException("原WPS进程仍在运行，已停止自动重启。请检查关闭提示或后台驻留状态。");}
                }
            }
            ticket.State="Starting";ticket.Save();
            // Start the exact host executable, not a filename association.
            int launchedPid;using(var launched=Process.Start(new ProcessStartInfo(ticket.Executable){UseShellExecute=false,WorkingDirectory=Path.GetDirectoryName(ticket.Executable)})) {if(launched==null)throw new IOException("未能启动 WPS。");launchedPid=launched.Id;}
            var deadline=DateTime.UtcNow.AddSeconds(60);while(DateTime.UtcNow<deadline) {
                object candidate=null;try {var clsid=new Guid(
#if WPS_WORD
                    "000209FF-0000-4b30-A977-D214852036FF"
#else
                    "45540001-5750-5300-4B49-4E47534F4655"
#endif
                );GetActiveObject(ref clsid,IntPtr.Zero,out candidate);int pid=Pid(candidate);
                    if(pid==launchedPid&&!ticket.ExistingProcesses.Contains(pid))using(var fresh=Process.GetProcessById(pid))if(fresh.SessionId==ticket.SessionId&&fresh.StartTime.ToUniversalTime().Ticks>ticket.Started&&string.Equals(Path.GetDirectoryName(fresh.MainModule.FileName),Path.GetDirectoryName(ticket.Executable),StringComparison.OrdinalIgnoreCase)){app=candidate;candidate=null;break;}
                }catch{}finally{if(candidate!=null)Marshal.ReleaseComObject(candidate);}Thread.Sleep(200);
            }
            if(app==null)throw new IOException("未取得新的WPS进程，可能被路由到其他已有实例；恢复清单已保留。");
            Application.EnableVisualStyles();((dynamic)app).Visible=true;
            // Programmatic Open must retain normal trust prompts rather than
            // inheriting automation's permissive macro setting.
            ((dynamic)app).AutomationSecurity=2; // msoAutomationSecurityByUI
            while(true) {ticket.State="Reopening";ticket.Save();object books=
#if WPS_WORD
                ((dynamic)app).Documents;
#else
                ((dynamic)app).Workbooks;
#endif
            try {foreach(var entry in ticket.Books.Where(b=>!b.Reopened)) {
                object workbook=null;try {if(!WpsShared.SourcePaths.Exists(entry.Path))throw new FileNotFoundException("文件已不存在。");workbook=((dynamic)books).Open(entry.Path,
#if WPS_WORD
                    false,
#else
                    0,
#endif
                    entry.ReadOnly);
                    if(workbook==null||!string.Equals(Path.GetFullPath(Convert.ToString(((dynamic)workbook).FullName)),Path.GetFullPath(entry.Path),StringComparison.OrdinalIgnoreCase)||Pid(app)==ticket.ProcessId)throw new IOException("重新打开的文件身份不匹配。");
                    entry.Reopened=true;entry.Error=null;}
                catch(Exception error){entry.Error=error.GetType().Name+" HRESULT="+error.HResult.ToString("X8");}finally{if(workbook!=null)Marshal.ReleaseComObject(workbook);ticket.Save();}
            }
            var active=ticket.Books.FirstOrDefault(b=>b.Active&&b.Reopened);if(active!=null)for(int i=1;i<=Convert.ToInt32(((dynamic)books).Count);i++){object workbook=((dynamic)books).Item(i);try{if(string.Equals(Convert.ToString(((dynamic)workbook).FullName),active.Path,StringComparison.OrdinalIgnoreCase)){((dynamic)workbook).Activate();break;}}finally{Marshal.ReleaseComObject(workbook);}}
            }finally{Marshal.ReleaseComObject(books);}
            ticket.State=ticket.Books.All(b=>b.Reopened)?"Completed":"PartiallyReopened";ticket.Save();
            if(ticket.State=="Completed")break;
            string message="以下文件未能重新打开：\r\n"+string.Join("\r\n",ticket.Books.Where(b=>!b.Reopened).Select(b=>b.Path+"："+b.Error))+"\r\n可处理文件问题后点击重试；已打开文件不会重复打开。";
            if(MessageBox.Show(message,"重新打开文件",MessageBoxButtons.RetryCancel,MessageBoxIcon.Warning)!=DialogResult.Retry)return 1;
            }
            return 0;
        }catch(Exception error){if(ticket!=null){ticket.Error=error.Message;ticket.State=ticket.State=="PartiallyReopened"?ticket.State:"Failed";try{ticket.Save();}catch{}}Application.EnableVisualStyles();MessageBox.Show(error.Message+(ticket==null?"":"\r\n恢复清单："+RestartTicket.FileFor(ticket.Id)),"WPS重启未完成",MessageBoxButtons.OK,MessageBoxIcon.Warning);return 1;}
        finally{if(app!=null)Marshal.ReleaseComObject(app);}
    }
}
