using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → Agent-API → 国外Agent → Gemini
// 控件：btnAgentGemini
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace WordUpdater
{
    public partial class Connect
    {
        public void OnOpenUrl_btnAgentGemini(object control) { OpenAgentUrl("btnAgentGemini"); }

        public IPictureDisp GetImage_btnAgentGemini(object control) { return GetAgentImage("btnAgentGemini"); }
    }
}
