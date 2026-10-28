using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 操作 → 自定配置
// 控件：menuAgentCustomization
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuAgentCustomization(object control) { return GetAgentImage("menuAgentCustomization"); }
    }
}
