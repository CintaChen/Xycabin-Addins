using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → Agent-API → 国内Agent → DeepSeek
// 控件：btnAgentDeepSeek
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace WordUpdater
{
    public partial class Connect
    {
        public void OnOpenUrl_btnAgentDeepSeek(object control) { OpenAgentUrl("btnAgentDeepSeek"); }

        public IPictureDisp GetImage_btnAgentDeepSeek(object control) { return GetAgentImage("btnAgentDeepSeek"); }
    }
}
