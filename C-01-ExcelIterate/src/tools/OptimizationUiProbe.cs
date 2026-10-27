using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Excel_Iterate;

public static class OptimizationUiProbe
{
    private const BindingFlags All=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Static|BindingFlags.Instance;
    [DllImport("user32.dll",CharSet=CharSet.Unicode,SetLastError=true)] private static extern IntPtr CreateDesktop(string name,IntPtr device,IntPtr mode,uint flags,uint access,IntPtr attributes);
    [DllImport("user32.dll",SetLastError=true)] private static extern bool SetThreadDesktop(IntPtr desktop);
    [DllImport("user32.dll")] private static extern bool CloseDesktop(IntPtr desktop);
    private static void Layout(Control control){control.CreateControl();control.PerformLayout();foreach(Control child in control.Controls)Layout(child);}
    [STAThread] public static int Main(string[] args)
    {
        IntPtr desktop=CreateDesktop("optimization-ui-"+Guid.NewGuid().ToString("N"),IntPtr.Zero,IntPtr.Zero,0,0x000F01FF,IntPtr.Zero);
        if(desktop==IntPtr.Zero)throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        Exception failure=null;
        try
        {
            var worker=new Thread(delegate(){try{if(!SetThreadDesktop(desktop))throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());Render(args);}catch(Exception error){failure=error;}});
            worker.SetApartmentState(ApartmentState.STA);worker.Start();worker.Join();
            if(failure!=null){Console.WriteLine("FAILED: "+failure);return 1;}return 0;
        }
        finally {CloseDesktop(desktop);}
    }
    private static void Render(string[] args)
    {
        string directory=Path.GetFullPath(args[0]);Directory.CreateDirectory(directory);Application.EnableVisualStyles();
        var assembly=typeof(Connect).Assembly;
        assembly.GetType("Excel_Iterate.DialogSizeStore").GetField("RootOverride",All).SetValue(null,Path.Combine(directory,"private-settings"));
        assembly.GetType("Excel_Iterate.FormatSession").GetField("SourceOverride",All).SetValue(null,Path.Combine(directory,"missing-settings.xml"));
        var factory=assembly.GetType("Excel_Iterate.DialogPreviewFactory").GetMethod("Create",All);
        var render=Assembly.LoadFrom(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"UiPreviewProbe.exe")).GetType("WpsUiChecks.UiPreviewProbe");
        foreach(string id in new[]{"elimination.receivable","elimination.revenuecost"})
        foreach(int percent in new[]{100,150,200})
        foreach(bool custom in new[]{false,true})
        using(var form=(Form)factory.Invoke(null,new object[]{id,new Size(1180,700)}))
        {
            form.Show();Application.DoEvents();
            if(form.WindowState!=FormWindowState.Normal)throw new Exception("Size preview must never maximize.");
            if(custom)
            {
                form.Controls.Find("sourceWorkbook",true)[0].Width=560;
                form.Controls.Find("sourceSheetBox",true)[0].Width=240;
            }
            if(percent!=100){form.Scale(new SizeF(percent/100f,percent/100f));render.GetMethod("ScaleFonts",All).Invoke(null,new object[]{form,percent/100f});}
            for(int pass=0;pass<3;pass++){Layout(form);Application.DoEvents();}
            foreach(string rowName in new[]{"sourceRow","targetRow"})
            {
                var row=form.Controls.Find(rowName,true)[0];
                if(row.Bottom>row.Parent.ClientSize.Height)
                    throw new Exception(rowName+" container is clipped at "+percent+"%: "+row.Bounds+" / "+row.Parent.ClientSize);
                foreach(Control field in row.Controls)
                    if(field.Right>row.ClientSize.Width||field.Bottom>row.ClientSize.Height)
                        throw new Exception(rowName+" field is clipped at "+percent+"%: "+field.Bounds+" / "+row.ClientSize);
            }
            using(var bitmap=new Bitmap(form.Width,form.Height))
            {form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,form.Size));bitmap.Save(Path.Combine(directory,id+"-"+percent+(custom?"-custom":"")+".png"),ImageFormat.Png);}
            Console.WriteLine("PASS private-desktop production preview "+id+" "+percent+"% "+(custom?"custom":"default")+" control scaling; client="+form.ClientSize+" actual device DPI="+form.DeviceDpi);
        }
        var production=assembly.GetType("Excel_Iterate.EliminationForm");
        foreach(var kind in new[]{EliminationKind.Receivable,EliminationKind.RevenueCost})
        using(var form=(Form)Activator.CreateInstance(production,All,null,new object[]{kind},null))
        {
            form.Show();Application.DoEvents();
            if(form.WindowState!=FormWindowState.Maximized)throw new Exception("Production form must initially maximize.");
            form.WindowState=FormWindowState.Normal;Application.DoEvents();Layout(form);
            if(form.WindowState!=FormWindowState.Normal||form.ClientSize.Width<=0)throw new Exception("Restoring production form failed.");
            form.WindowState=FormWindowState.Maximized;Application.DoEvents();Layout(form);
            if(form.WindowState!=FormWindowState.Maximized)throw new Exception("Re-maximizing production form failed.");
            Console.WriteLine("PASS private-desktop production "+kind+" initial maximize, restore, re-maximize; actual device DPI="+form.DeviceDpi);
        }
    }
}
