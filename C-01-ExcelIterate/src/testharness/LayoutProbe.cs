using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;
using Excel_Iterate;

// =====================================================================
// 布局测量探针（临时诊断工具）：经反射实例化 internal 的「合并抵消」弹窗，
// 输出来源/目标卡片及其各层子控件的实际像素高度，
// 用于定位「卡片下方留白」的真正来源。不属于交付物。
// =====================================================================
internal static class LayoutProbe
{
    static void Main()
    {
        Assembly asm = Assembly.LoadFrom(AppDomain.CurrentDomain.BaseDirectory + "Excel_Iterate.dll");
        Type formType = asm.GetType("Excel_Iterate.EliminationForm");
        if (formType == null) { Console.WriteLine("EliminationForm 未找到"); return; }
        Form form = (Form)Activator.CreateInstance(formType,
            BindingFlags.Instance | BindingFlags.NonPublic, null,
            new object[] { EliminationKind.RevenueCost }, null);
        form.Show();
        Application.DoEvents();
        form.PerformLayout();
        Application.DoEvents();

        Console.WriteLine("DPI=" + form.DeviceDpi + " ClientSize=" + form.ClientSize + " WindowState=" + form.WindowState);
        TableLayoutPanel root = FindRoot(form);
        if (root == null) { Console.WriteLine("root 未找到"); return; }
        Console.WriteLine("root.Padding=" + root.Padding);
        var heights = root.GetRowHeights();
        for (int r = 0; r < root.RowCount; r++)
        {
            Control ctl = root.GetControlFromPosition(0, r);
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                "  row {0}: style={1}({2}) actual={3} control={4}",
                r, root.RowStyles[r].SizeType, root.RowStyles[r].Height,
                heights[r], ctl == null ? "-" : ctl.GetType().Name + " H=" + ctl.Height));
        }

        Console.WriteLine("--- sourceCard 层级 ---");
        Control sourceCard = root.GetControlFromPosition(0, 0);
        Dump("sourceCard", sourceCard);
        TableLayoutPanel sourceRow = FirstTlp(sourceCard);
        if (sourceRow != null)
        {
            Dump("  sourceRow", sourceRow);
            foreach (Control c in sourceRow.Controls)
            {
                Dump("    " + c.GetType().Name, c);
                FlowLayoutPanel flow = c as FlowLayoutPanel;
                if (flow != null)
                    foreach (Control inner in flow.Controls)
                        Dump("      " + inner.GetType().Name + "[" + inner.Name + "]", inner);
            }
        }

        Console.WriteLine("--- targetCard 层级 ---");
        Control targetCard = root.GetControlFromPosition(0, 3);
        Dump("targetCard", targetCard);
        TableLayoutPanel targetRow = FirstTlp(targetCard);
        if (targetRow != null)
        {
            Dump("  targetRow", targetRow);
            foreach (Control c in targetRow.Controls)
            {
                Dump("    " + c.GetType().Name, c);
                FlowLayoutPanel flow = c as FlowLayoutPanel;
                if (flow != null)
                    foreach (Control inner in flow.Controls)
                        Dump("      " + inner.GetType().Name, inner);
            }
        }

        Console.WriteLine("--- 总结 ---");
        Console.WriteLine("sourceCard.Height=" + sourceCard.Height + " targetCard.Height=" + targetCard.Height);
        try { form.Close(); } catch { }
    }

    static TableLayoutPanel FirstTlp(Control parent)
    {
        foreach (Control c in parent.Controls)
        {
            var tlp = c as TableLayoutPanel;
            if (tlp != null) return tlp;
        }
        return null;
    }

    static TableLayoutPanel FindRoot(Form form)
    {
        foreach (Control c in form.Controls)
        {
            var tlp = c as TableLayoutPanel;
            if (tlp != null && tlp.RowCount == 12) return tlp;
        }
        return null;
    }

    static void Dump(string label, Control c)
    {
        Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
            "{0}: H={1} PreferredH={2} Margin={3} Padding={4} Dock={5} AutoSize={6}",
            label, c.Height, c.GetPreferredSize(Size.Empty).Height,
            c.Margin, c.Padding, c.Dock, c.AutoSize));
    }
}
