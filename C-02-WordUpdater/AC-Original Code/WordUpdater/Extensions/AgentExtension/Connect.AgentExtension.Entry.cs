using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：WordUpdater → 扩展 → Agent扩展
// 控件：btnAgentExtension
// 共用实现：@shared/src/addin/Connect.Agent.cs
namespace WordUpdater
{
    public partial class Connect
    {
        public void OnToggleAgentExtension(object control, bool pressed)
        {
            string warning = AgentController.SetVisible(pressed);
            if (warning.Length != 0) UiDialog.Show(warning, "Agent扩展", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public IPictureDisp GetImage_btnAgentExtension(object control) { return GetAgentImage("btnAgentExtension"); }
    }
}
