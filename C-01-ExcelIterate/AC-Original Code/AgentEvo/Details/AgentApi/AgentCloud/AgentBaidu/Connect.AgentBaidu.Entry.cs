using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → 智能云-API → 百度智能云
// 控件：btnAgentBaidu
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnOpenUrl_btnAgentBaidu(object control) { OpenAgentUrl("btnAgentBaidu"); }

        public IPictureDisp GetImage_btnAgentBaidu(object control) { return GetAgentImage("btnAgentBaidu"); }
    }
}
