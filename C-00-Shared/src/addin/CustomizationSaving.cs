using System;
using System.Windows.Forms;
#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    internal static class CustomizationSaving
    {
        internal static Action SavedForProbe;
        internal static void AfterSave(Form source){
            LiveCustomization.Notify();
            if(SavedForProbe!=null){SavedForProbe();source.Close();return;}
            bool menu=source is CommandCustomizationForm;
            UiDialog.Show(source,menu?"已保存并生效。":"已保存，下次打开对应弹窗时生效。","自定配置",MessageBoxButtons.OK,MessageBoxIcon.Information);source.Close();
        }
    }
}
