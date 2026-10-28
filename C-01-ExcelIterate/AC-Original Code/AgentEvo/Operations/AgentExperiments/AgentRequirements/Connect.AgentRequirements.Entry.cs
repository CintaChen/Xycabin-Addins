using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 操作 → 实验计划 → 需求记录
// 控件：btnAgentRequirements
// 共用实现：@shared/src/addin/Connect.Customization.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnAgentRequirements(object control) {OpenCustomization("notes");}

        public IPictureDisp GetImage_btnAgentRequirements(object control) {return GetAgentImage("btnAgentRequirements");}
    }
}
