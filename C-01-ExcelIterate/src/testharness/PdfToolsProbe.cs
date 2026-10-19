using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

// Offline production-form and embedded-worker test. Does not connect to Office/WPS.
internal static class PdfToolsProbe
{
    private const BindingFlags Any = BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            Application.EnableVisualStyles();
            Assembly assembly = Assembly.LoadFrom(Path.GetFullPath(args[0]));
            string output = Path.GetFullPath(args[1]); Directory.CreateDirectory(output);
            string ns = assembly.GetName().Name == "WordUpdater" ? "WordUpdater" : "Excel_Iterate";
            Type factory = assembly.GetType(ns + ".DialogPreviewFactory", true);
            foreach (string mode in new[] { "current", "convert", "watermark", "split", "merge" })
            {
                string id = "file.pdf." + mode;
                using (var form = (Form)factory.GetMethod("CreateProgramDefault", Any).Invoke(null, new object[] { id }))
                {
                    if (!form.ShowInTaskbar) throw new Exception("Unexpected tool window style: " + id);
                    form.StartPosition = FormStartPosition.Manual; form.Location = new Point(-20000, -20000);
                    form.Show(); Application.DoEvents(); form.PerformLayout();
                    foreach (Size size in new[] { new Size(960, 700), new Size(800, 600) })
                    {
                        form.ClientSize = size; form.PerformLayout(); Application.DoEvents();
                        using (var image = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(image, new Rectangle(Point.Empty, image.Size)); image.Save(Path.Combine(output, mode + "-" + size.Width + ".png")); }
                    }
                }
            }
            // Exercise extraction and the real child process, including dependency redirects.
            string input = Path.Combine(output, "bridge image.png");
            using (var bitmap = new Bitmap(80, 120)) { using (var g = Graphics.FromImage(bitmap)) g.Clear(Color.CornflowerBlue); bitmap.Save(input); }
            Type requestType = assembly.GetType(ns + ".PdfToolRequest", true);
            object request = Activator.CreateInstance(requestType, true);
            requestType.GetField("Mode", Any).SetValue(request, "convert");
            requestType.GetField("Output", Any).SetValue(request, output);
            requestType.GetField("Files", Any).SetValue(request, new[] { input });
            string report = (string)assembly.GetType(ns + ".PdfWorkerBridge", true).GetMethod("Run", Any).Invoke(null, new object[] { request, new Func<bool>(() => false) });
            if (!Directory.GetFiles(output, "bridge image*.pdf").Any()) throw new Exception(report);
            Console.WriteLine("PASS: five production previews at default/minimum sizes; embedded worker extraction and image conversion.");
            return 0;
        }
        catch (Exception error) { Console.WriteLine(error); return 1; }
    }
}
