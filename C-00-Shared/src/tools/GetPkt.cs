using System;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

class GetPkt
{
    [DataContract]
    private sealed class Identity
    {
        [DataMember(Name="name")] public string Name;
        [DataMember(Name="version")] public string Version;
        [DataMember(Name="language")] public string Language;
        [DataMember(Name="processorArchitecture")] public string Architecture;
        [DataMember(Name="publicKeyToken",EmitDefaultValue=false)] public string Token;
        internal static Identity Read(AssemblyName name)
        {
            byte[] token = name.GetPublicKeyToken();
            return new Identity { Name = name.Name, Version = name.Version.ToString(),
                Language = string.IsNullOrEmpty(name.CultureName) ? "neutral" : name.CultureName,
                Architecture = name.ProcessorArchitecture.ToString().ToLowerInvariant(),
                Token = token == null || token.Length == 0 ? null : BitConverter.ToString(token).Replace("-", "").ToLowerInvariant() };
        }
    }
    [DataContract]
    private sealed class Metadata
    {
        [DataMember(Name="path")] public string Path;
        [DataMember(Name="identity")] public Identity Identity;
        [DataMember(Name="references")] public Identity[] References;
    }
    static int Main(string[] args)
    {
        if (args.Length < 1) { Console.WriteLine("usage: GetPkt <dll>"); return 2; }
        try
        {
            Console.OutputEncoding = new UTF8Encoding(false);
            if (args[0] == "--json")
            {
                if (args.Length < 2) return 2;
                var metadata = new Metadata[args.Length - 1];
                for (int i = 1; i < args.Length; i++)
                {
                    string path = Path.GetFullPath(args[i]);
                    // Reflection-only loading reads metadata and never runs DLL code.
                    AssemblyName[] references = Assembly.ReflectionOnlyLoadFrom(path).GetReferencedAssemblies();
                    var identities = new Identity[references.Length];
                    for (int j = 0; j < references.Length; j++) identities[j] = Identity.Read(references[j]);
                    metadata[i - 1] = new Metadata { Path = path, Identity = Identity.Read(AssemblyName.GetAssemblyName(path)), References = identities };
                }
                using (var stream = new MemoryStream())
                {
                    new DataContractJsonSerializer(typeof(Metadata[])).WriteObject(stream, metadata);
                    Console.WriteLine(Encoding.UTF8.GetString(stream.ToArray()));
                }
                return 0;
            }
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
