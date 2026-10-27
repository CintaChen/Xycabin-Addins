using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Threading;
using System.Windows.Forms;
using WpsShared;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    internal sealed class RestartComBag:IDisposable {
        private readonly List<object> items=new List<object>();internal object Own(object value){if(value!=null)items.Add(value);return value;}
        public void Dispose(){foreach(var value in items.AsEnumerable().Reverse())try{if(Marshal.IsComObject(value))Marshal.ReleaseComObject(value);}catch{}items.Clear();}
    }
    internal static class RestartProduct {
#if WPS_WORD
        internal const string Name="WordUpdater",Host="WPS 文字",Files="文档";
#else
        internal const string Name="Excel_Iterate",Host="WPS 表格",Files="工作簿";
#endif
    }
    internal interface IRestartHost {
        List<RestartBook> Capture();
        void Save(RestartBook book);
        void Verify(List<RestartBook> books);
        void Quit();
    }
    internal interface IRestartGate:IDisposable {void Prepare(List<RestartBook> books);void Commit();void Abort();}
    internal static class RestartWorkflow {
        internal static void Run(IRestartHost host,IRestartGate gate) {
            // Nothing closes until every save and the external helper handshake pass.
            var books=host.Capture();foreach(var book in books)host.Save(book);host.Verify(books);
            try {gate.Prepare(books);gate.Commit();host.Quit();}catch {gate.Abort();throw;}
        }
    }
    internal static class CustomizationRestart
    {
        private static Control dispatcher;private static object host;private static IntPtr owner;private static bool busy;
        [DllImport("user32.dll")]internal static extern uint GetWindowThreadProcessId(IntPtr handle,out uint process);
        internal static int HostPid(object app) {
            IntPtr handle=IntPtr.Zero;try{handle=new IntPtr(Convert.ToInt64(((dynamic)app).Hwnd));}catch{}
            if(handle==IntPtr.Zero)using(var bag=new RestartComBag()){object window=bag.Own(((dynamic)app).ActiveWindow);handle=new IntPtr(Convert.ToInt64(((dynamic)window).Hwnd));}
            uint pid;GetWindowThreadProcessId(handle,out pid);if(pid==0)throw new InvalidOperationException("无法核对当前宿主实例，未重启。");return (int)pid;
        }
        internal static void Bind(object app,IntPtr handle) {
            if(app==null||HostPid(app)!=Process.GetCurrentProcess().Id)throw new InvalidOperationException("当前宿主对象不属于此加载项进程，未执行自动重启。");
            host=app;owner=handle;if(dispatcher==null){dispatcher=new Control();var ignored=dispatcher.Handle;}
            CustomizationSaving.RequestRestart=delegate(Form source) {
                if(source.InvokeRequired)source.BeginInvoke((MethodInvoker)delegate{source.Close();});else source.Close();
                dispatcher.BeginInvoke((MethodInvoker)Run);
            };
        }
        private sealed class Owner:IWin32Window {internal IntPtr Value;public IntPtr Handle{get{return Value;}}}
        private static void Run() {
            if(busy)return;busy=true;
            using(var guard=new Mutex(false,"Local\\"+RestartProduct.Name+".Restart."+Process.GetCurrentProcess().Id+"."+Process.GetCurrentProcess().StartTime.Ticks)) {
                bool held=false;try {
                    try{held=guard.WaitOne(0);}catch(AbandonedMutexException){held=true;}if(!held)return;
                    string reason=Connect.RestartBlockReason();if(reason!=null)throw new InvalidOperationException(reason);
                    Confirm();
                    using(var adapter=new ComRestartHost(host))using(var gate=new RestartHelperGate())RestartWorkflow.Run(adapter,gate);
                }catch(Exception ex){Connect.Log("保存重启停止："+ex.GetType().Name+" HRESULT="+ex.HResult.ToString("X8"));UiDialog.Show(new Owner{Value=owner},"配置已保存，本次重启未完成：\r\n"+ex.Message+"\r\n已成功保存的文件保留保存结果；设置将在完全退出并重开 WPS 后生效。","保存与重启",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
                finally{if(held)guard.ReleaseMutex();busy=false;}
            }
        }
        private static void Confirm() {
            using(var dialog=new Form {Text="保存与重启 "+RestartProduct.Host,ControlBox=false,ShowInTaskbar=false,StartPosition=FormStartPosition.CenterParent,FormBorderStyle=FormBorderStyle.FixedDialog,ClientSize=new Size(540,170)}) {
                var message=Ui.Label("配置已保存。点击确认后，将自动保存当前 "+RestartProduct.Host+" 实例的全部"+RestartProduct.Files+"，重启并重新打开这些文件。",11,false);message.Dock=DockStyle.Fill;message.Padding=new Padding(20);message.AutoSize=false;
                var button=Ui.Button("确认重启",130);Ui.StylePrimary(button);button.DialogResult=DialogResult.OK;var bar=new FlowLayoutPanel {Dock=DockStyle.Bottom,Height=55,Padding=new Padding(15,4,15,8),FlowDirection=FlowDirection.RightToLeft};bar.Controls.Add(button);dialog.Controls.Add(message);dialog.Controls.Add(bar);dialog.AcceptButton=button;
                dialog.FormClosing+=delegate(object sender,FormClosingEventArgs e){if(e.CloseReason==CloseReason.UserClosing&&dialog.DialogResult!=DialogResult.OK)e.Cancel=true;};
                if(dialog.ShowDialog(new Owner{Value=owner})!=DialogResult.OK)throw new OperationCanceledException("尚未确认重启，未关闭文件。");
            }
        }
        internal static void Release() {CustomizationSaving.RequestRestart=null;host=null;if(dispatcher!=null){dispatcher.Dispose();dispatcher=null;}}
    }
    internal sealed class ComRestartHost:IRestartHost,IDisposable
    {
        private readonly object app;private readonly RestartComBag bag=new RestartComBag();private readonly Dictionary<RestartBook,object> captured=new Dictionary<RestartBook,object>();
        internal ComRestartHost(object application){app=application;}
        private object Collection {get {
#if WPS_WORD
            return ((dynamic)app).Documents;
#else
            return ((dynamic)app).Workbooks;
#endif
        }}
        private object Active {get {
#if WPS_WORD
            return ((dynamic)app).ActiveDocument;
#else
            return ((dynamic)app).ActiveWorkbook;
#endif
        }}
        public List<RestartBook> Capture() {
            if(CustomizationRestart.HostPid(app)!=Process.GetCurrentProcess().Id)throw new InvalidOperationException("宿主实例已变化，未关闭文件。");
            #if WPS_WORD
            Connect.ClearPendingFlash();
#endif
            dynamic books=bag.Own(Collection);object active=bag.Own(Convert.ToInt32(books.Count)==0?null:Active);var list=new List<RestartBook>();
            for(int i=1;i<=Convert.ToInt32(books.Count);i++) {object value=bag.Own(books.Item(i));dynamic book=value;var entry=new RestartBook{Name=Convert.ToString(book.Name),Path=Convert.ToString(book.FullName),ReadOnly=Convert.ToBoolean(book.ReadOnly),Active=Same(value,active)};captured.Add(entry,value);list.Add(entry);}return list;
        }
        private static bool Same(object a,object b) {if(a==null||b==null)return a==b;IntPtr x=IntPtr.Zero,y=IntPtr.Zero;try{x=Marshal.GetIUnknownForObject(a);y=Marshal.GetIUnknownForObject(b);return x==y;}finally{if(x!=IntPtr.Zero)Marshal.Release(x);if(y!=IntPtr.Zero)Marshal.Release(y);}}
        public void Save(RestartBook entry) {
            dynamic book=captured[entry];string path=Convert.ToString(book.Path);bool saved=Convert.ToBoolean(book.Saved);
            bool local=Path.IsPathRooted(entry.Path)&&File.Exists(entry.Path);
            if(path.Length==0||!local||(entry.ReadOnly&&!saved)) {
                book.Activate();object dialogs=bag.Own(((dynamic)app).Dialogs);
#if WPS_WORD
                object dialog=bag.Own(((dynamic)dialogs).Item(84));
#else
                object dialog=bag.Own(((dynamic)dialogs).Item(5));
#endif
                object result=((dynamic)dialog).Show();if(result!=null&&
#if WPS_WORD
                    Convert.ToInt32(result)!=-1
#else
                    !Convert.ToBoolean(result)
#endif
                )throw new OperationCanceledException(entry.Name+"：已取消另存为，未关闭文件。");
            }else if(!saved){
#if !WPS_WORD
                Connect.PrepareWorkbookSave(entry.Path);
#endif
                book.Save();}
            entry.Path=Convert.ToString(book.FullName);entry.ReadOnly=Convert.ToBoolean(book.ReadOnly);
            if(!Convert.ToBoolean(book.Saved)||Convert.ToString(book.Path).Length==0||!Path.IsPathRooted(entry.Path)||!File.Exists(entry.Path))throw new IOException(entry.Name+"：保存未完成或没有可重新打开的本地文件。");
        }
        public void Verify(List<RestartBook> entries) {
            dynamic books=bag.Own(Collection);if(Convert.ToInt32(books.Count)!=entries.Count)throw new IOException("保存期间文件集合已变化，请重试。");
            var actual=new List<object>();for(int i=1;i<=entries.Count;i++)actual.Add(bag.Own(books.Item(i)));
            foreach(var entry in entries)if(!actual.Any(b=>Same(b,captured[entry]))||!Convert.ToBoolean(((dynamic)captured[entry]).Saved)||!string.Equals(entry.Path,Convert.ToString(((dynamic)captured[entry]).FullName),StringComparison.OrdinalIgnoreCase))throw new IOException(entry.Name+"：身份、路径或保存状态已变化，未关闭文件。");
        }
        public void Quit(){Connect.CloseAgentExtension();((dynamic)app).Quit();}
        public void Dispose(){bag.Dispose();}
    }
    internal sealed class RestartHelperGate:IRestartGate
    {
        private RestartTicket ticket;private EventWaitHandle ready,go,abort;private Process helper;
        public void Prepare(List<RestartBook> books) {
            var process=Process.GetCurrentProcess();ticket=new RestartTicket {Id=Guid.NewGuid().ToString("N"),User=WindowsIdentity.GetCurrent().User.Value,ProcessId=process.Id,SessionId=process.SessionId,Started=process.StartTime.ToUniversalTime().Ticks,Executable=process.MainModule.FileName,Books=books};
            ticket.ExistingProcesses=Process.GetProcessesByName("et").Concat(Process.GetProcessesByName("wps")).Select(p=>{try{return p.Id;}finally{p.Dispose();}}).ToList();ticket.Save();
            string exe=ReleaseHelper();var acl=new System.Security.AccessControl.EventWaitHandleSecurity();acl.AddAccessRule(new System.Security.AccessControl.EventWaitHandleAccessRule(WindowsIdentity.GetCurrent().User,System.Security.AccessControl.EventWaitHandleRights.FullControl,System.Security.AccessControl.AccessControlType.Allow));bool created;
            ready=new EventWaitHandle(false,EventResetMode.ManualReset,"Local\\Wps.Restart."+ticket.Id+".ready",out created,acl);go=new EventWaitHandle(false,EventResetMode.ManualReset,"Local\\Wps.Restart."+ticket.Id+".go",out created,acl);abort=new EventWaitHandle(false,EventResetMode.ManualReset,"Local\\Wps.Restart."+ticket.Id+".abort",out created,acl);
            helper=Process.Start(new ProcessStartInfo(exe,ticket.Id){UseShellExecute=false,CreateNoWindow=true,WorkingDirectory=Path.GetDirectoryName(exe)});
            if(helper==null||!ready.WaitOne(10000))throw new IOException("重启辅助程序未就绪，未关闭文件。恢复记录："+RestartTicket.FileFor(ticket.Id));
        }
        private static string ReleaseHelper() {
            byte[] bytes;using(var source=typeof(Connect).Assembly.GetManifestResourceStream("WpsShared.RestartHelper.exe"))using(var memory=new MemoryStream()){if(source==null)throw new IOException("缺少重启辅助程序。");source.CopyTo(memory);bytes=memory.ToArray();}
            string hash;using(var sha=SHA256.Create())hash=BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-","");string folder=Path.Combine(RestartTicket.Root,"Helper",hash);Directory.CreateDirectory(folder);string file=Path.Combine(folder,"WpsRestartHelper.exe");
            if(File.Exists(file)){using(var sha=SHA256.Create())if(sha.ComputeHash(File.ReadAllBytes(file)).SequenceEqual(sha.ComputeHash(bytes)))return file;throw new IOException("重启辅助程序摘要不一致，未关闭文件。");}
            string temp=file+"."+Guid.NewGuid().ToString("N")+".tmp";try{File.WriteAllBytes(temp,bytes);try{File.Move(temp,file);}catch(IOException){if(!File.Exists(file)||!File.ReadAllBytes(file).SequenceEqual(bytes))throw;}}finally{if(File.Exists(temp))File.Delete(temp);}return file;
        }
        public void Commit(){ticket.State="WaitingForExit";ticket.Save();go.Set();}
        public void Abort(){if(abort!=null)abort.Set();if(ticket!=null){ticket.State="Stopped";try{ticket.Save();}catch{}}}
        public void Dispose(){if(ready!=null)ready.Dispose();if(go!=null)go.Dispose();if(abort!=null)abort.Dispose();if(helper!=null)helper.Dispose();}
    }
}
