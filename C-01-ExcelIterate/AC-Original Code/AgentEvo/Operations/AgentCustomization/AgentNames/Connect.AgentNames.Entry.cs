using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 操作 → 自定配置 → 自定名称
// 控件：btnAgentNames
// 共用实现：@shared/src/addin/Connect.Customization.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnAgentNames(object control) {OpenCustomization("name");}

        public IPictureDisp GetImage_btnAgentNames(object control) {return GetAgentImage("btnAgentNames");}
    }
}
