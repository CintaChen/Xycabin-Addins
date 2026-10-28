using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 功能：WordUpdater → 操作 → 文件操作
// 控件：menuFileOperations
namespace WordUpdater
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuFileOperations(object control) { return GetButtonImageCore("menuFileOperations"); }
    }
}
