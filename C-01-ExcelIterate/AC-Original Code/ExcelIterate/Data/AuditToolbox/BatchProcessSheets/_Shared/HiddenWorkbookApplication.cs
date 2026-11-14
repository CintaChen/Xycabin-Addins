using System;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    // Owns a newly-created, independent automation process, never a user host.
    internal sealed class HiddenWorkbookApplication : IDisposable
    {
        [DllImport("user32.dll")]private static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
        private readonly ComBag references=new ComBag();
        internal object Application {get;private set;}
        private bool owns;
        internal HiddenWorkbookApplication(object source)
        {
            object candidate=null;
            try
            {
                string name=Convert.ToString(((dynamic)source).Name);string progId=name.IndexOf("Excel",StringComparison.OrdinalIgnoreCase)>=0?"Excel.Application":"ket.Application";
                Type type=Type.GetTypeFromProgID(progId);if(type==null)throw new InvalidOperationException("无法创建独立后台表格进程。");
                candidate=references.Own(Activator.CreateInstance(type));uint original,worker;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)source).Hwnd)),out original);
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)candidate).Hwnd)),out worker);
                if(original==0||worker==0||original==worker)throw new InvalidOperationException("宿主未提供独立后台进程，已停止，避免隐藏或关闭你的操作窗口。");
                using(var bag=new ComBag()){dynamic books=bag.Own(((dynamic)candidate).Workbooks);if(Convert.ToBoolean(((dynamic)candidate).Visible)||Convert.ToInt32(books.Count)!=0)throw new InvalidOperationException("后台进程已含工作簿，未接管或关闭该进程。");}
                owns=true;Application=candidate;
                dynamic app=candidate;app.Visible=false;app.DisplayAlerts=false;app.EnableEvents=false;app.ScreenUpdating=false;app.AskToUpdateLinks=false;app.AutomationSecurity=3;
            }
            catch{try{if(owns&&candidate!=null)((dynamic)candidate).Quit();}finally{owns=false;Application=null;references.Dispose();}throw;}
        }
        public void Dispose()
        {try{if(owns&&Application!=null)((dynamic)Application).Quit();}finally{owns=false;Application=null;references.Dispose();}}
    }
}
