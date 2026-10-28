using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → Agent-API
// 控件：menuAgentSites
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuAgentSites(object control) { return GetAgentImage("menuAgentSites"); }
    }
}
