using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

// 功能：Excel-Iterate → 数据 → 审计工具箱 → 合并抵消 → 往来抵消
// 控件：btnEliminationReceivable
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnEliminationReceivable(object control)
        { OpenElimination(EliminationKind.Receivable); }

        public IPictureDisp GetImage_btnEliminationReceivable(object control)
        { return GetButtonImageCore("btnEliminationReceivable"); }
    }
}
