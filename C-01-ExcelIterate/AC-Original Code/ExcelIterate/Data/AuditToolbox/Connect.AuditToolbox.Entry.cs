using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

// 功能：Excel-Iterate → 数据 → 审计工具箱
// 控件：menuAuditToolbox
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuAuditToolbox(object control)
        { return GetButtonImageCore("menuAuditToolbox"); }
    }
}
