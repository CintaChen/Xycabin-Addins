using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Excel;

namespace Excel_Iterate
{
    // =====================================================================
    // VSTO 加载项入口（WPS 12.1 经 vstor 加载，挂在 WPS 表格 ET）
    //
    // 契约（从本机 CPAHelper VSTO 加载项逆向确认）：
    //   [StartupObject(0)] + 继承 Microsoft.Office.Tools.AddInBase
    //   ctor: (Microsoft.Office.Tools.Excel.ApplicationFactory, IServiceProvider)
    //   基类构造: (factory, serviceProvider, "ThisAddIn", "ThisAddIn")
    //
    // 加载链：WPS 表格读 HKCU\...\Excel\Addins\<Name> 的 Manifest 值
    //   -> vstor 弹“自定义项安装程序” -> 用户点安装
    //   -> 加载本程序集 -> 找 [StartupObject] 的 AddInBase 子类
    //   -> new ThisAddIn(factory, sp) -> BeginInit/Initialize/FinishInitialization
    //   -> CreateRibbonExtensibilityObject() -> Ribbon XML 回调
    // =====================================================================
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObject(0)]
    [System.Security.Permissions.PermissionSet(
        System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
    public class ThisAddIn : AddInBase
    {
        private Connect _ribbon;

        public ThisAddIn(ApplicationFactory factory, IServiceProvider serviceProvider)
            : base(factory, serviceProvider, "ThisAddIn", "ThisAddIn")
        {
        }

        protected override void Initialize()
        {
            base.Initialize();
            Connect.Log("VSTO Initialize");
        }

        protected override void FinishInitialization()
        {
            base.FinishInitialization();
            HostDialogService.Register();
            Connect.Log("VSTO FinishInitialization");
            // 尝试提前建立与 WPS Application 的连接（失败则在点击时重试）
            Connect.EnsureApplication();
            Connect.StartHostConnection();
        }

        protected override void OnShutdown()
        {
            Connect.CloseAgentExtension();
            Connect.Log("VSTO OnShutdown");
            Connect.StopHostConnection();
            LiveCustomization.Stop();
            Connect.CloseTrackerForm();       // 关掉还开着的「增强追踪」非模态窗
            Connect.CloseMergeForm();         // 关掉还开着的「表头匹配合并」非模态窗
            Connect.CloseDataMigrationForm();
            Connect.CloseEliminationForm();   // 关掉还开着的「合并抵消」非模态窗
            Connect.ClosePdfTools();
            Connect.HostApplication = null;
            base.OnShutdown();
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Connect.Log("CreateRibbonExtensibilityObject 被调用");
            if (_ribbon == null)
            {
                _ribbon = new Connect();
            }
            return _ribbon;
        }
    }
}
