using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

// 功能：Excel-Iterate → 数据 → 审计工具箱 → 合并抵消 → 收入成本抵消
// 控件：btnEliminationRevenueCost
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnEliminationRevenueCost(object control)
        { OpenElimination(EliminationKind.RevenueCost); }

        public IPictureDisp GetImage_btnEliminationRevenueCost(object control)
        { return GetButtonImageCore("btnEliminationRevenueCost"); }
    }
}
