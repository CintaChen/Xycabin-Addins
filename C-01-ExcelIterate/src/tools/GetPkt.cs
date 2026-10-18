using System;
using System.Reflection;

class GetPkt
{
    static int Main(string[] args)
    {
        if (args.Length < 1) { Console.WriteLine("usage: GetPkt <dll>"); return 2; }
        try
        {
            AssemblyName n = AssemblyName.GetAssemblyName(args[0]);
            byte[] t = n.GetPublicKeyToken();
            string token = (t == null || t.Length == 0) ? "NO-STRONG-NAME" : BitConverter.ToString(t).Replace("-", "").ToLowerInvariant();
            Console.WriteLine(n.Name + " | version=" + n.Version + " | pkt=" + token);
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("ERROR: " + ex.GetType().Name + ": " + ex.Message);
            return 1;
        }
    }
}
