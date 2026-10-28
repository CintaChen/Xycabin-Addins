using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 详情 → 调用API → Agent-API → 国内Agent
// 控件：menuAgentChina
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuAgentChina(object control) { return GetAgentImage("menuAgentChina"); }
    }
}
