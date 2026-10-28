using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → 智能云-API
// 控件：menuAgentCloud
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuAgentCloud(object control) { return GetAgentImage("menuAgentCloud"); }
    }
}
