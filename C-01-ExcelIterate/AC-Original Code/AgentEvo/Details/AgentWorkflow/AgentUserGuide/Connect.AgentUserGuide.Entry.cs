using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 开发流程 → 用户预览
// 控件：btnAgentUserGuide
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnOpenAgentUserGuide(object control) { OpenAgentDocument(true); }

        public IPictureDisp GetImage_btnAgentUserGuide(object control) { return GetAgentImage("btnAgentUserGuide"); }
    }
}
