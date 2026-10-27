using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;
using WpsShared;

internal static class InstallFlowProbe
{
    private static int checks;

    private static void Check(bool value, string name)
    {
        checks++;
        if (!value) throw new Exception("FAIL: " + name);
    }

    private static InstallResult Success()
    {
        return new InstallResult(InstallStatus.Success, "fake success", 0);
    }

    private static void VerifyInstall(string runtime, int code, InstallResult cleanup,
        InstallStatus expected, string expectedOrder)
    {
        List<string> calls = new List<string>();
        InstallResult result = InstallFlow.Install(new InstallActions {
            FindRuntime = delegate { return runtime; },
            CleanupSubscriptions = delegate(string tool) { calls.Add("cleanup"); return cleanup; },
            ReleasePayload = delegate { calls.Add("payload"); },
            InvokeInstall = delegate(string tool) { calls.Add("install"); return code; },
            Register = delegate { calls.Add("register"); },
            Complete = delegate { calls.Add("complete"); }
        });
        Check(result.Status == expected, "install result " + expected);
        Check(string.Join(",", calls.ToArray()) == expectedOrder, "install side effects " + expected);
        if (expected == InstallStatus.MissingRuntime)
            Check(!result.ExitCode.HasValue, "missing runtime does not report official completion");
        if (code != 0 && expectedOrder.IndexOf("install", StringComparison.Ordinal) >= 0)
            Check(result.ExitCode == code, "official failure/cancel code preserved");
    }

    private static void VerifyUninstall(string runtime, InstallResult cleanup,
        InstallStatus expected, string expectedOrder)
    {
        List<string> calls = new List<string>();
        InstallResult result = InstallFlow.Uninstall(new UninstallActions {
            FindRuntime = delegate { return runtime; },
            CleanupSubscriptions = delegate(string tool) { calls.Add("cleanup"); return cleanup; },
            RemovePayloadAndRegistration = delegate { calls.Add("delete"); }
        });
        Check(result.Status == expected, "uninstall result " + expected);
        Check(string.Join(",", calls.ToArray()) == expectedOrder, "uninstall retains files until official cleanup succeeds");
    }

    private static int Main()
    {
        try
        {
            string custom = @"D:\Runtime Tools\VSTOInstaller.exe";
            List<RegistryView> queried = new List<RegistryView>();
            string found = RuntimeLocator.Find(delegate(RegistryView view) {
                queried.Add(view); return view == RegistryView.Registry32 ? custom : null;
            }, new string[0], delegate(string path) { return path == custom; });
            Check(found == custom, "custom InstallerPath is accepted");
            Check(queried.Contains(RegistryView.Registry64) && queried.Contains(RegistryView.Registry32), "both registry views queried");
            string folder = @"D:\Other Runtime";
            found = RuntimeLocator.Find(delegate(RegistryView view) { return folder; }, new string[0],
                delegate(string path) { return path == Path.Combine(folder, "VSTOInstaller.exe"); });
            Check(found == Path.Combine(folder, "VSTOInstaller.exe"), "InstallerPath folder form");
            string common = @"C:\Program Files (x86)\Common Files";
            string commonTool = Path.Combine(common, @"Microsoft Shared\VSTO\10.0", "VSTOInstaller.exe");
            found = RuntimeLocator.Find(delegate(RegistryView view) { return @"D:\missing\VSTOInstaller.exe"; },
                new string[] { @"C:\Program Files\Common Files", common }, delegate(string path) { return path == commonTool; });
            Check(found == commonTool, "common x86 path fallback after stale registry");
            string obsolete = Path.Combine(common, @"Microsoft Shared\VSTO\9.0", "VSTOInstaller.exe");
            Check(RuntimeLocator.Find(delegate(RegistryView view) { return null; }, new string[] { common },
                delegate(string path) { return path == obsolete; }) == null, "obsolete pre-.NET4 default runtime is rejected");
            Check(RuntimeLocator.Find(delegate(RegistryView view) { return custom; }, new string[] { common },
                delegate(string path) { return false; }) == null, "runtime requires existing tool");
            Check(RuntimeLocator.Find(delegate(RegistryView view) { return @"D:\wrong.exe"; }, new string[0],
                delegate(string path) { return true; }) == null, "wrong InstallerPath executable rejected");
            Check(RuntimeLocator.Find(delegate(RegistryView view) { return "VSTOInstaller.exe"; }, new string[0],
                delegate(string path) { return true; }) == null, "relative InstallerPath rejected");

            VerifyInstall(null, 0, Success(), InstallStatus.MissingRuntime, "");
            VerifyInstall(custom, 0, Success(), InstallStatus.Success, "cleanup,payload,install,register,complete");
            VerifyInstall(custom, -500, Success(), InstallStatus.Cancelled, "cleanup,payload,install");
            VerifyInstall(custom, -300, Success(), InstallStatus.Failed, "cleanup,payload,install");
            VerifyInstall(custom, -400, Success(), InstallStatus.Failed, "cleanup,payload,install");
            VerifyInstall(custom, 0, InstallResult.FromExitCode(-300, "cleanup"), InstallStatus.Failed, "cleanup");
            VerifyInstall(custom, 0, InstallResult.FromExitCode(-500, "cleanup"), InstallStatus.Cancelled, "cleanup");
            VerifyUninstall(null, Success(), InstallStatus.MissingRuntime, "");
            VerifyUninstall(custom, Success(), InstallStatus.Success, "cleanup,delete");
            VerifyUninstall(custom, InstallResult.FromExitCode(-300, "cleanup"), InstallStatus.Failed, "cleanup");

            string owned = "WordUpdater.vsto";
            Check(RuntimeLocator.IsOwnedUrl("file:///D:/插件/WordUpdater.vsto", owned), "Chinese file path is owned");
            Check(RuntimeLocator.IsOwnedUrl("file:///D:/A%20B/WordUpdater.vsto|vstolocal", owned), "escaped path and registry suffix are owned");
            Check(RuntimeLocator.IsOwnedUrl("file:///D:/plugins/wordupdater.vsto", owned), "Windows case insensitive ownership");
            Check(!RuntimeLocator.IsOwnedUrl("file:///D:/plugins/WordUpdater.vsto.backup", owned), "similarly named foreign manifest preserved");
            Check(!RuntimeLocator.IsOwnedUrl("file:///D:/WordUpdater.vsto/Foreign.vsto", owned), "foreign manifest under similarly named directory preserved");
            Check(!RuntimeLocator.IsOwnedUrl("file:///D:/plugins/CPAHelper.vsto", owned), "third party subscription preserved");
            Check(!RuntimeLocator.IsOwnedUrl("https://example.com/WordUpdater.vsto", owned), "remote subscription preserved");
            Check(!RuntimeLocator.IsOwnedUrl("file:///D:/plugins/WordUpdater.vsto#foreign", owned), "fragment subscription preserved");
            Check(!RuntimeLocator.IsOwnedUrl("file:///D:/plugins/WordUpdater.vsto?foreign=1", owned), "query subscription preserved");
            Check(RuntimeLocator.IsOwnedUrl("file:///D:/A%23B/%E6%8F%92%E4%BB%B6/WordUpdater.vsto", owned), "escaped hash and Chinese path owned");
            Check(!RuntimeLocator.IsOwnedUrl(null, owned), "null subscription preserved");

            List<string> removed = new List<string>();
            InstallResult subscriptions = SubscriptionCleanup.Remove(new string[] {
                "file:///D:/plugins/CPAHelper.vsto", "file:///D:/plugins/WordUpdater.vsto.backup",
                "file:///D:/A%20B/WordUpdater.vsto", "file:///D:/A B/WordUpdater.vsto",
                "file:///D:/plugins/WordUpdater.vsto#foreign"
            }, owned, delegate(string url) { removed.Add(url); return 0; }, delegate(string text) { });
            Check(subscriptions.Status == InstallStatus.Success && removed.Count == 1,
                "fake launcher receives only exact owned URL, canonical duplicates removed");
            removed.Clear();
            subscriptions = SubscriptionCleanup.Remove(new string[] {
                "file:///D:/first/WordUpdater.vsto", "file:///D:/second/WordUpdater.vsto"
            }, owned, delegate(string url) { removed.Add(url); return -300; }, delegate(string text) { });
            Check(subscriptions.Status == InstallStatus.Failed && removed.Count == 1,
                "unsubscribe failure stops before further subscriptions");

            InstallResult exception = InstallFlow.Install(new InstallActions {
                FindRuntime = delegate { return custom; }, CleanupSubscriptions = delegate(string tool) { return Success(); },
                ReleasePayload = delegate { throw new IOException("fake locked file"); }
            });
            Check(exception.Status == InstallStatus.Failed && exception.Message.Contains("fake locked file"), "payload failure does not report success");
            Console.WriteLine("InstallFlowProbe PASS: " + checks + " checks (fake runtime, no installation or registry changes)");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex.Message); return 1; }
    }
}
