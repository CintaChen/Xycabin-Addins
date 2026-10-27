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
        public void OnAgentSort(object control) {OpenCustomization("sort");}
        public void OnAgentNames(object control) {OpenCustomization("name");}
        public void OnAgentFormats(object control) {OpenCustomization("format");}
        public void OnAgentRequirements(object control) {OpenCustomization("notes");}
        public void OnAgentOcrPlaceholder(object control) {}
        public IPictureDisp GetImage_btnAgentSort(object control) {return GetAgentImage("btnAgentSort");}
        public IPictureDisp GetImage_btnAgentNames(object control) {return GetAgentImage("btnAgentNames");}
        public IPictureDisp GetImage_btnAgentFormats(object control) {return GetAgentImage("btnAgentFormats");}
        public IPictureDisp GetImage_menuAgentExperiments(object control) {return GetAgentImage("menuAgentExperiments");}
        public IPictureDisp GetImage_btnAgentOcrPlaceholder(object control) {return GetAgentImage("btnAgentOcrPlaceholder");}
        public IPictureDisp GetImage_btnAgentRequirements(object control) {return GetAgentImage("btnAgentRequirements");}
    }
}
