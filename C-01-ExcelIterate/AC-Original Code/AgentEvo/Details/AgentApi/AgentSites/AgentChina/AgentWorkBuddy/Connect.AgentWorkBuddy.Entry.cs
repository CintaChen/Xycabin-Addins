using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → Agent-API → 国内Agent → WorkBuddy
// 控件：btnAgentWorkBuddy
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnOpenUrl_btnAgentWorkBuddy(object control) { OpenAgentUrl("btnAgentWorkBuddy"); }

        public IPictureDisp GetImage_btnAgentWorkBuddy(object control) { return GetAgentImage("btnAgentWorkBuddy"); }
    }
}
