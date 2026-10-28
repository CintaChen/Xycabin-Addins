using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：Agent-Evo → 操作 → 实验计划 → OCR识别
// 控件：btnAgentOcrPlaceholder
// 共用实现：@shared/src/addin/Connect.Customization.cs
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnAgentOcrPlaceholder(object control) {}

        public IPictureDisp GetImage_btnAgentOcrPlaceholder(object control) {return GetAgentImage("btnAgentOcrPlaceholder");}
    }
}
