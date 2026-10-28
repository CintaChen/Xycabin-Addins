using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 开发流程 → 开发预览
// 控件：btnAgentReadme
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnOpenAgentReadme(object control) { OpenAgentDocument(false); }

        public IPictureDisp GetImage_btnAgentReadme(object control) { return GetAgentImage("btnAgentReadme"); }
    }
}
