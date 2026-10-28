using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → Agent-API → 国内Agent → TRAE
// 控件：btnAgentTrae
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace WordUpdater
{
    public partial class Connect
    {
        public void OnOpenUrl_btnAgentTrae(object control) { OpenAgentUrl("btnAgentTrae"); }

        public IPictureDisp GetImage_btnAgentTrae(object control) { return GetAgentImage("btnAgentTrae"); }
    }
}
