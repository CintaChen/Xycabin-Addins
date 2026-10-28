using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 开发流程
// 控件：menuAgentWorkflow
namespace WordUpdater
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuAgentWorkflow(object control) { return GetAgentImage("menuAgentWorkflow"); }
    }
}
