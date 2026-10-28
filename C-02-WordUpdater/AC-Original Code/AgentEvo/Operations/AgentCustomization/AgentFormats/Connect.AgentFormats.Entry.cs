using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 操作 → 自定配置 → 自定格式
// 控件：btnAgentFormats
// 共用实现：@shared/src/addin/Connect.Customization.cs
namespace WordUpdater
{
    public partial class Connect
    {
        public void OnAgentFormats(object control) {OpenCustomization("format");}

        public IPictureDisp GetImage_btnAgentFormats(object control) {return GetAgentImage("btnAgentFormats");}
    }
}
