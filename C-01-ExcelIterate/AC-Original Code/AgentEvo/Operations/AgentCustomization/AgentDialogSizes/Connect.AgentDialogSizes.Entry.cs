using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 操作 → 自定配置 → 交互弹窗
// 控件：btnAgentDialogSizes
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnOpenAgentDialogSizes(object control)
        { OnOpenAgentDialogSizesCore(control); }

        public IPictureDisp GetImage_btnAgentDialogSizes(object control) { return GetAgentImage("btnAgentDialogSizes"); }
    }
}
