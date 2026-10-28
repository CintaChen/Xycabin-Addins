using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → Agent-API → 国内Agent → 豆包
// 控件：btnAgentDoubao
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnOpenUrl_btnAgentDoubao(object control) { OpenAgentUrl("btnAgentDoubao"); }

        public IPictureDisp GetImage_btnAgentDoubao(object control) { return GetAgentImage("btnAgentDoubao"); }
    }
}
