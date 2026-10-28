using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → 智能云-API → 阿里智能云
// 控件：btnAgentAliyun
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnOpenUrl_btnAgentAliyun(object control) { OpenAgentUrl("btnAgentAliyun"); }

        public IPictureDisp GetImage_btnAgentAliyun(object control) { return GetAgentImage("btnAgentAliyun"); }
    }
}
