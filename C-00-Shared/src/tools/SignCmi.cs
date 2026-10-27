// SHA-256 ClickOnce/VSTO signing. Certificate selection is explicit and read-only.
// SignCmi <manifest> <thumbprint> [expected-subject]
// SignCmi --check-certificate <thumbprint> <expected-subject>
using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Xml;

internal static class SignCmi
{
    private const string SD = "System.Deployment, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
    private const BindingFlags ALLI = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
    private const BindingFlags ALLS = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

    private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        try { return Run(args); }
        catch (Exception ex)
        {
            while (ex.InnerException != null) ex = ex.InnerException;
            Console.WriteLine("FAIL " + ex.GetType().Name + ": " + ex.Message);
            return 1;
        }
    }

    private static int Run(string[] args)
    {
        if (args.Length < 2 || args.Length > 3)
            throw new ArgumentException("usage: SignCmi <manifest|--check-certificate> <thumbprint> [expected-subject]");
        bool certificateOnly = args[0] == "--check-certificate";
        if (certificateOnly && args.Length != 3)
            throw new ArgumentException("Certificate checks require the expected subject.");
        string thumb = args[1].Trim().ToUpperInvariant();
        if (thumb.Length != 40 || !System.Text.RegularExpressions.Regex.IsMatch(thumb, "^[0-9A-F]{40}$"))
            throw new ArgumentException("Invalid certificate thumbprint.");
        string path = certificateOnly ? null : Path.GetFullPath(args[0]);
        if (!certificateOnly && !WpsShared.SourcePaths.Exists(path))
            throw new FileNotFoundException("Manifest is missing.", path);
        using (X509Certificate2 cert = FindCertificate(thumb))
        {
            if (args.Length == 3 && !string.Equals(cert.Subject, args[2], StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Certificate subject does not match the project.");
            DateTime now = DateTime.UtcNow;
            if (now < cert.NotBefore.ToUniversalTime() || now > cert.NotAfter.ToUniversalTime())
                throw new InvalidOperationException("Certificate is outside its validity period.");
            using (RSA rsa = cert.GetRSAPrivateKey())
            {
                if (rsa == null) throw new InvalidOperationException("Certificate private key is unavailable.");
                // A read-only signing challenge checks access to the private key, not just its metadata.
                byte[] data = Encoding.UTF8.GetBytes("WPS_CERTIFICATE_CHECK");
                byte[] signature = rsa.SignData(data, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
                if (!rsa.VerifyData(data, signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1))
                    throw new CryptographicException("Certificate private-key check failed.");
                if (certificateOnly)
                {
                    Console.WriteLine("CERTIFICATE_OK " + cert.Thumbprint.ToUpperInvariant());
                    return 0;
                }

                Assembly asm = Assembly.Load(SD);
                Type signerType = asm.GetType("System.Deployment.Internal.CodeSigning.CmiManifestSigner2", true);
                Type manifestType = asm.GetType("System.Deployment.Internal.CodeSigning.SignedCmiManifest2", true);
                object signer = Activator.CreateInstance(signerType, ALLI, null,
                    new object[] { (AsymmetricAlgorithm)rsa, cert, true }, null);
                XmlDocument doc = new XmlDocument();
                doc.PreserveWhitespace = true;
                doc.XmlResolver = null;
                using (XmlReader reader = XmlReader.Create(path,
                    new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
                    doc.Load(reader);
                MethodInfo remove = manifestType.GetMethod("RemoveExistingSignature", ALLS);
                if (remove != null) remove.Invoke(null, new object[] { doc });
                object signed = Activator.CreateInstance(manifestType, ALLI, null, new object[] { doc, true }, null);
                MethodInfo sign = manifestType.GetMethod("Sign", ALLI, null, new Type[] { signerType }, null);
                if (sign == null) throw new MissingMethodException("SHA-256 ClickOnce signer is unavailable.");
                sign.Invoke(signed, new object[] { signer });
                doc.Save(path);
                Console.WriteLine("SIGNED_SHA256 " + cert.Thumbprint.ToUpperInvariant() + " " + path);
                return 0;
            }
        }
    }

    private static X509Certificate2 FindCertificate(string thumb)
    {
        foreach (StoreLocation location in new StoreLocation[] { StoreLocation.CurrentUser, StoreLocation.LocalMachine })
        {
            using (X509Store store = new X509Store(StoreName.My, location))
            {
                try { store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly); }
                catch (CryptographicException) { continue; }
                foreach (X509Certificate2 candidate in store.Certificates)
                    if (string.Equals(candidate.Thumbprint, thumb, StringComparison.OrdinalIgnoreCase))
                        return new X509Certificate2(candidate);
            }
        }
        throw new InvalidOperationException("Pinned certificate was not found: " + thumb);
    }
}
