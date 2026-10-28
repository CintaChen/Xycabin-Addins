using System;
using System.Windows.Forms;
#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    public partial class Connect
    {
        internal static IntPtr CustomizationOwner(){return GetHostWindowHandle();}
        internal static void ShowCustomizationWindow(Form form) {IntPtr owner=GetHostWindowHandle();if(owner==IntPtr.Zero)form.Show();else form.Show(new WindowWrapper(owner));}
        private static void OpenCustomization(string mode) {try {
if(!EnsureApplication())throw new InvalidOperationException("请先打开对应 WPS 文件。");
            CustomizationWindows.Open(mode);}catch(Exception ex){UiDialog.Show(ex.Message,"自定配置",MessageBoxButtons.OK,MessageBoxIcon.Warning);}}

    }
}
