using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 开发流程 → XyCabin
// 控件：btnAgentHome
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace WordUpdater
{
    public partial class Connect
    {
        public void OnAgentHome(object control)
        { OnAgentHomeCore(control); }

        public IPictureDisp GetImage_btnAgentHome(object control) { return GetAgentImage("btnAgentHome"); }
    }
}
