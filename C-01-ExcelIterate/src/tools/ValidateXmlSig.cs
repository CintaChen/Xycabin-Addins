// Read-only SHA-256 VSTO verification: both signatures, pinned certificate,
// license-to-manifest identity/hash binding, and certificate/public-key consistency.
// Authenticode signs r:license, so its SignedXml scope must be that subtree.
// Certificate trust is granted separately by the standard interactive VSTO installer.
// ValidateXmlSig [--thumbprint <fingerprint>] <manifest> [manifest2 ...]
using System;
using System.Reflection;
using System.Text;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography.X509Certificates;
using System.Xml;

public class RSAPKCS1SHA256SignatureDescription : SignatureDescription
{
    public RSAPKCS1SHA256SignatureDescription()
    {
        KeyAlgorithm = typeof(RSACryptoServiceProvider).FullName;
        DigestAlgorithm = typeof(SHA256Managed).FullName;
        FormatterAlgorithm = typeof(RSAPKCS1SignatureFormatter).FullName;
        DeformatterAlgorithm = typeof(RSAPKCS1SignatureDeformatter).FullName;
    }
    public override HashAlgorithm CreateDigest() { return SHA256.Create(); }
    public override AsymmetricSignatureDeformatter CreateDeformatter(AsymmetricAlgorithm key)
    {
        AsymmetricSignatureDeformatter result = base.CreateDeformatter(key);
        result.SetHashAlgorithm("SHA256");
        return result;
    }
    public override AsymmetricSignatureFormatter CreateFormatter(AsymmetricAlgorithm key)
    {
        AsymmetricSignatureFormatter result = base.CreateFormatter(key);
        result.SetHashAlgorithm("SHA256");
        return result;
    }
}

internal static class ValidateXmlSig
{
    private const string DS = "http://www.w3.org/2000/09/xmldsig#";
    private const string AS = "http://schemas.microsoft.com/windows/pki/2005/Authenticode";
    private const string REL = "urn:mpeg:mpeg21:2003:01-REL-R-NS";
    private const string C14N = "http://www.w3.org/2001/10/xml-exc-c14n#";
    private const BindingFlags ALLI = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

    private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        CryptoConfig.AddAlgorithm(typeof(RSAPKCS1SHA256SignatureDescription), DS + "rsa-sha256");
        CryptoConfig.AddAlgorithm(typeof(SHA256Managed), DS + "sha256", "http://www.w3.org/2001/04/xmlenc#sha256");
        string thumb = null;
        int first = 0;
        if (args.Length > 0 && args[0] == "--thumbprint")
        {
            if (args.Length < 3) { Console.WriteLine("FAIL Missing fingerprint or manifest."); return 2; }
            thumb = args[1].Trim().ToUpperInvariant();
            if (!System.Text.RegularExpressions.Regex.IsMatch(thumb, "^[0-9A-F]{40}$"))
            { Console.WriteLine("FAIL Invalid fingerprint."); return 2; }
            first = 2;
        }
        if (first == args.Length) { Console.WriteLine("usage: ValidateXmlSig [--thumbprint <fingerprint>] <manifest> [...]"); return 2; }
        int failures = 0;
        for (int i = first; i < args.Length; i++)
        {
            try
            {
                VerifyManifest(args[i], thumb);
                Console.WriteLine("PASS SHA256_DOUBLE_SIGNATURE " + args[i]);
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null) ex = ex.InnerException;
                Console.WriteLine("FAIL " + args[i] + ": " + ex.GetType().Name + ": " + ex.Message);
                failures++;
            }
        }
        return failures == 0 ? 0 : 1;
    }

    private static XmlNamespaceManager Namespaces(XmlDocument doc)
    {
        XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("asm", "urn:schemas-microsoft-com:asm.v1");
        ns.AddNamespace("asmv2", "urn:schemas-microsoft-com:asm.v2");
        ns.AddNamespace("ds", DS);
        ns.AddNamespace("msrel", "http://schemas.microsoft.com/windows/rel/2005/reldata");
        ns.AddNamespace("r", REL);
        ns.AddNamespace("as", AS);
        return ns;
    }

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new CryptographicException(message);
    }

    private static void VerifyManifest(string path, string thumb)
    {
        XmlDocument doc = new XmlDocument();
        doc.PreserveWhitespace = true;
        doc.XmlResolver = null;
        using (XmlReader reader = XmlReader.Create(path,
            new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
            doc.Load(reader);
        XmlNamespaceManager ns = Namespaces(doc);
        Require(doc.GetElementsByTagName("Signature", DS).Count == 2, "Exactly two signatures are required.");
        XmlNodeList outer = doc.DocumentElement.SelectNodes("ds:Signature", ns);
        Require(outer.Count == 1, "Exactly one outer StrongName signature is required.");
        XmlElement strong = (XmlElement)outer[0];
        Require(strong.GetAttribute("Id") == "StrongNameSignature", "StrongName signature identity is missing.");
        XmlElement license = (XmlElement)strong.SelectSingleNode("ds:KeyInfo/msrel:RelData/r:license", ns);
        Require(license != null && doc.GetElementsByTagName("license", REL).Count == 1, "Exactly one signed license is required.");
        XmlElement auth = (XmlElement)license.SelectSingleNode("r:issuer/ds:Signature", ns);
        Require(auth != null && auth.GetAttribute("Id") == "AuthenticodeSignature", "Authenticode signature is missing.");
        XmlNodeList certificateNodes = auth.SelectNodes("ds:KeyInfo/ds:X509Data/ds:X509Certificate", ns);
        Require(certificateNodes.Count == 1 && doc.GetElementsByTagName("X509Certificate", DS).Count == 1, "Exactly one signing certificate is required.");
        using (X509Certificate2 cert = new X509Certificate2(Convert.FromBase64String(certificateNodes[0].InnerText)))
        using (RSA rsa = cert.GetRSAPublicKey())
        {
            Require(rsa != null, "An RSA certificate is required.");
            Require(thumb == null || string.Equals(cert.Thumbprint, thumb, StringComparison.OrdinalIgnoreCase),
                "Signing certificate differs from Configure_program.py.");
            DateTime now = DateTime.UtcNow;
            Require(now >= cert.NotBefore.ToUniversalTime() && now <= cert.NotAfter.ToUniversalTime(), "Signing certificate is outside its validity period.");
            XmlElement publisher = (XmlElement)doc.DocumentElement.SelectSingleNode("asmv2:publisherIdentity", ns);
            XmlNode publisherSubject = license.SelectSingleNode("r:grant/as:AuthenticodePublisher/as:X509SubjectName", ns);
            Require(publisher != null && publisher.GetAttribute("name") == cert.Subject
                && publisherSubject != null && publisherSubject.InnerText == cert.Subject, "Publisher and certificate subjects differ.");
            RSAParameters key = rsa.ExportParameters(false);
            foreach (XmlElement signature in new XmlElement[] { strong, auth })
            {
                XmlNode modulus = signature.SelectSingleNode("ds:KeyInfo/ds:KeyValue/ds:RSAKeyValue/ds:Modulus", ns);
                XmlNode exponent = signature.SelectSingleNode("ds:KeyInfo/ds:KeyValue/ds:RSAKeyValue/ds:Exponent", ns);
                Require(modulus != null && exponent != null
                    && modulus.InnerText == Convert.ToBase64String(key.Modulus)
                    && exponent.InnerText == Convert.ToBase64String(key.Exponent), "Certificate and signature public keys differ.");
            }
            VerifySignature(doc, strong, rsa);
            // Import the license as the document root, reproducing the Authenticode signing scope.
            XmlDocument licenseDoc = new XmlDocument();
            licenseDoc.PreserveWhitespace = true;
            licenseDoc.XmlResolver = null;
            licenseDoc.AppendChild(licenseDoc.ImportNode(license, true));
            XmlElement licenseSignature = (XmlElement)licenseDoc.GetElementsByTagName("Signature", DS)[0];
            VerifySignature(licenseDoc, licenseSignature, rsa);
        }
        VerifyCmiBinding(doc, ns);
    }

    private static void VerifySignature(XmlDocument scope, XmlElement signature, RSA key)
    {
        SignedXml signed = new SignedXml(scope);
        signed.LoadXml(signature);
        Require(signed.SignatureMethod == DS + "rsa-sha256", "SHA-256 RSA signing is required.");
        Require(signed.SignedInfo.CanonicalizationMethod == C14N, "Unexpected signature canonicalization.");
        Require(signed.SignedInfo.References.Count == 1, "Exactly one local signature reference is required.");
        foreach (Reference reference in signed.SignedInfo.References)
        {
            Require(reference.Uri == "" && reference.DigestMethod == DS + "sha256", "Unexpected signature reference/digest.");
            Require(reference.TransformChain.Count == 2
                && reference.TransformChain[0].Algorithm == DS + "enveloped-signature"
                && reference.TransformChain[1].Algorithm == C14N, "Unexpected signature transforms.");
            foreach (Transform transform in reference.TransformChain)
                if (transform is XmlDsigEnvelopedSignatureTransform) transform.Context = signature;
        }
        Require(signed.CheckSignature(key), "Invalid " + signature.GetAttribute("Id") + " signature.");
    }

    private static void VerifyCmiBinding(XmlDocument doc, XmlNamespaceManager ns)
    {
        // Use the same .NET 4.8 ClickOnce implementation as SignCmi for token, identity and license hash.
        // Full Verify also checks certificate-chain trust. The VSTO installer separately grants
        // application trust to a self-signed add-in; that decision does not install a trusted root.
        Assembly assembly = Assembly.Load("System.Deployment, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a");
        Type manifestType = assembly.GetType("System.Deployment.Internal.CodeSigning.SignedCmiManifest2", true);
        Type flagsType = assembly.GetType("System.Deployment.Internal.CodeSigning.CmiManifestVerifyFlags", true);
        Type algorithmType = assembly.GetType("System.Deployment.Internal.CodeSigning.SupportedHashAlgorithm", true);
        object manifest = Activator.CreateInstance(manifestType, ALLI, null, new object[] { doc, true }, null);
        manifestType.GetMethod("Verify", ALLI).Invoke(manifest, new object[] { Enum.Parse(flagsType, "StrongNameOnly") });
        object sha256 = Enum.Parse(algorithmType, "SHA256");
        foreach (string method in new string[] { "VerifyAssemblyIdentity", "VerifyHash" })
            manifestType.GetMethod(method, ALLI).Invoke(manifest, new object[] { ns, sha256 });
    }
}
