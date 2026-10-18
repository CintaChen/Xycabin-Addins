// TrustAdd —— 把 VSTO 加载项加入 ClickOnce 包含列表（Inclusion List）。
//
// 这是同事机器上「WPS 不弹出 Microsoft Office 自定义项安装程序」的根因所在：
//   VSTO 遇到未知证书的自定义项时，本该弹对话框征求用户信任；但 WPS 的 VSTO
//   兼容层弹不出这个框，于是直接抛 SecurityException 静默失败：
//     SecurityException: 因为没有对此应用程序中的自定义功能授予信任……
//     用于签署部署清单的证书未知，自定义项本身(...)不在包含列表中。
//   区域 MyComputer。
// 而同目录其他能正常加载的插件（CPAHelper、审计助手），其公钥都登记在
//   HKCU\Software\Microsoft\VSTO\Security\Inclusion\<guid>
//      Url       = file:///<vsto 路径>
//      PublicKey = <RSAKeyValue><Modulus>..</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>
// 里，所以它们不需要询问就能加载。
//
// 本工具把 **部署清单里内嵌的签名证书**（不需要私钥）取出，按同样格式登记。
// 用法: TrustAdd.exe <vsto路径>
using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Xml;
using Microsoft.Win32;

internal static class TrustAdd
{
    internal const string InclusionKey = @"Software\Microsoft\VSTO\Security\Inclusion";

    /// <summary>从部署清单的签名里取出 X509 证书（不含私钥）。</summary>
    internal static X509Certificate2 ReadSigningCertificate(string vstoPath)
    {
        XmlDocument doc = new XmlDocument();
        doc.PreserveWhitespace = true;
        using (XmlReader xr = XmlReader.Create(vstoPath,
            new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit }))
        {
            doc.Load(xr);
        }
        XmlNodeList nodes = doc.GetElementsByTagName(
            "X509Certificate", "http://www.w3.org/2000/09/xmldsig#");
        if (nodes.Count == 0) return null;
        string b64 = nodes[0].InnerText.Trim();
        return new X509Certificate2(Convert.FromBase64String(b64));
    }

    /// <summary>证书公钥的 RSAKeyValue 文本，与 VSTO 自己写入包含列表的格式一致。</summary>
    internal static string PublicKeyToXml(X509Certificate2 cert)
    {
        RSA rsa = cert.GetRSAPublicKey();
        if (rsa == null) return null;
        RSAParameters p = rsa.ExportParameters(false);
        return "<RSAKeyValue><Modulus>"
            + Convert.ToBase64String(p.Modulus)
            + "</Modulus><Exponent>"
            + Convert.ToBase64String(p.Exponent)
            + "</Exponent></RSAKeyValue>";
    }

    /// <summary>登记包含列表条目。返回写入的子键名；已存在则原样返回。</summary>
    internal static string AddInclusion(string vstoUrl, X509Certificate2 cert)
    {
        using (RegistryKey root = Registry.CurrentUser.CreateSubKey(InclusionKey))
        {
            if (root == null) throw new InvalidOperationException("无法创建包含列表注册表项");
            string pubkey = PublicKeyToXml(cert);
            if (pubkey == null) throw new InvalidOperationException("无法读取证书公钥");

            // 幂等：同一个 Url + 同一把公钥已登记过就不要重复写
            foreach (string name in root.GetSubKeyNames())
            {
                using (RegistryKey k = root.OpenSubKey(name))
                {
                    if (k == null) continue;
                    object u = k.GetValue("Url");
                    object pk = k.GetValue("PublicKey");
                    if (u != null && pk != null &&
                        string.Equals(u.ToString(), vstoUrl, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(pk.ToString(), pubkey, StringComparison.Ordinal))
                    {
                        return name;
                    }
                }
            }

            string guid = Guid.NewGuid().ToString();
            using (RegistryKey k = root.CreateSubKey(guid))
            {
                if (k == null) throw new InvalidOperationException("无法写入包含列表条目");
                k.SetValue("Url", vstoUrl, RegistryValueKind.String);
                k.SetValue("PublicKey", pubkey, RegistryValueKind.String);
            }
            return guid;
        }
    }

    private static int Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("usage: TrustAdd <vsto路径>");
            return 2;
        }
        string path = Path.GetFullPath(args[0]);
        if (!File.Exists(path))
        {
            Console.WriteLine("找不到部署清单: " + path);
            return 3;
        }
        X509Certificate2 cert = ReadSigningCertificate(path);
        if (cert == null)
        {
            Console.WriteLine("部署清单里没有签名证书");
            return 4;
        }
        Console.WriteLine("  证书 = " + cert.Subject);
        Console.WriteLine("  指纹 = " + cert.Thumbprint);

        string url = "file:///" + path.Replace("\\", "/");
        Console.WriteLine("  Url  = " + url);
        string key = AddInclusion(url, cert);
        Console.WriteLine("  >>> 已登记包含列表: HKCU\\" + InclusionKey + "\\" + key);
        return 0;
    }
}
