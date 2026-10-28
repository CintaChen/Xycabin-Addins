using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → Agent-API → 国内Agent → Qoder
// 控件：btnAgentQoder
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace WordUpdater
{
    public partial class Connect
    {
        public void OnOpenUrl_btnAgentQoder(object control) { OpenAgentUrl("btnAgentQoder"); }

        public IPictureDisp GetImage_btnAgentQoder(object control) { return GetAgentImage("btnAgentQoder"); }
    }
}
