using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → Agent-API → 国外Agent → Claude
// 控件：btnAgentClaude
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnOpenUrl_btnAgentClaude(object control) { OpenAgentUrl("btnAgentClaude"); }

        public IPictureDisp GetImage_btnAgentClaude(object control) { return GetAgentImage("btnAgentClaude"); }
    }
}
