using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 操作 → 自定配置 → 自定排序
// 控件：btnAgentSort
// 共用实现：@shared/src/addin/Connect.Customization.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnAgentSort(object control) {OpenCustomization("sort");}

        public IPictureDisp GetImage_btnAgentSort(object control) {return GetAgentImage("btnAgentSort");}
    }
}
