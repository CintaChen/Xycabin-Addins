using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → Agent-API → 国内Agent → Kimi
// 控件：btnAgentKimi
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace WordUpdater
{
    public partial class Connect
    {
        public void OnOpenUrl_btnAgentKimi(object control) { OpenAgentUrl("btnAgentKimi"); }

        public IPictureDisp GetImage_btnAgentKimi(object control) { return GetAgentImage("btnAgentKimi"); }
    }
}
