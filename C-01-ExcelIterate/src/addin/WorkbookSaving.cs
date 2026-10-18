using System;
using System.IO;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal static class WorkbookSaving
    {
        internal static Func<string> ChooseNewPath;
        internal static string Save(object book, bool requireXlsx)
        {
            if (book == null) throw new InvalidOperationException("工作簿已关闭。");
            dynamic b = book;
            if (ComFormulaHost.Boolean(b.ReadOnly, "ReadOnly")) throw new InvalidOperationException("工作簿为只读，无法自动保存：" + b.FullName);
            string path = Convert.ToString(b.FullName);
            if (!Path.IsPathRooted(path))
            {
                string selected;
                if (ChooseNewPath != null) selected = ChooseNewPath();
                else selected = HostDialogService.SaveWorkbook(Path.GetFileNameWithoutExtension(path) + ".xlsx");
                if (string.IsNullOrWhiteSpace(selected)) throw new OperationCanceledException("已取消选择保存路径，本阶段未继续。");
                if (!string.Equals(Path.GetExtension(selected), ".xlsx", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("请选择 .xlsx 保存路径。");
                b.SaveAs(selected, 51); path = Convert.ToString(b.FullName);
            }
            else
            {
                if (requireXlsx && !string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("链接清单和替换目标需要 .xlsx 工作簿：" + path);
                if (!File.Exists(path) || !ComFormulaHost.Boolean(b.Saved, "Saved")) b.Save();
            }
            if (!File.Exists(path) || !ComFormulaHost.Boolean(b.Saved, "Saved")) throw new InvalidOperationException("自动保存未完成，工作簿仍有修改：" + path);
            if (requireXlsx && !string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("工作簿未保存为 .xlsx：" + path);
            return Path.GetFullPath(path);
        }
    }
}
