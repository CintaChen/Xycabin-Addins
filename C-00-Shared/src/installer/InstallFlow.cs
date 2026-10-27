using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;

namespace WpsShared
{
    internal enum InstallStatus { Success, Cancelled, MissingRuntime, Failed }

    internal sealed class InstallResult
    {
        internal readonly InstallStatus Status;
        internal readonly int? ExitCode;
        internal readonly string Message;

        internal InstallResult(InstallStatus status, string message, int? exitCode)
        {
            Status = status; Message = message; ExitCode = exitCode;
        }

        internal static InstallResult FromExitCode(int code, string operation)
        {
            if (code == 0) return new InstallResult(InstallStatus.Success, operation + "完成。", code);
            if (code == -500) return new InstallResult(InstallStatus.Cancelled, "已取消" + operation + "。", code);
            return new InstallResult(InstallStatus.Failed, operation + "失败，VSTOInstaller 返回 " + code + "。", code);
        }
    }

    // No UI or installation side effects: also used by SetupDoctor.
    internal static class RuntimeLocator
    {
        internal const string DownloadUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=105890";
        internal const string MissingMessage = "未找到 Microsoft Visual Studio Tools for Office Runtime（VSTO 运行库）。\r\n"
            + "微软官方下载地址：\r\n" + DownloadUrl + "\r\n\r\n"
            + "1. 点击「获取运行库」，或复制上述地址到浏览器，打开微软官方下载页。\r\n"
            + "2. 点击网页中的「Download（下载）」，下载并运行 vstor_redist.exe，按提示完成安装或修复（可能需要管理员权限）。\r\n"
            + "3. 运行库安装完成后，重新运行本加载项安装包，再重试当前操作。\r\n\r\n"
            + "本次操作未释放加载项文件、修改加载项登记或清理已有安装。";

        internal static string FindInstalledTool()
        {
            return Find(ReadInstallerPath, new string[] {
                Environment.GetEnvironmentVariable("CommonProgramFiles"),
                Environment.GetEnvironmentVariable("CommonProgramFiles(x86)"),
                Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86)
            }, File.Exists);
        }

        private static string ReadInstallerPath(RegistryView view)
        {
            try
            {
                using (RegistryKey machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                using (RegistryKey key = machine.OpenSubKey(@"SOFTWARE\Microsoft\VSTO Runtime Setup\v4"))
                    return key == null ? null : key.GetValue("InstallerPath") as string;
            }
            catch { return null; }
        }

        internal static string Find(Func<RegistryView, string> readInstallerPath,
            IEnumerable<string> commonRoots, Func<string, bool> fileExists)
        {
            foreach (RegistryView view in new RegistryView[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                string candidate = readInstallerPath(view);
                if (string.IsNullOrWhiteSpace(candidate)) continue;
                candidate = Environment.ExpandEnvironmentVariables(candidate.Trim().Trim('"'));
                if (!Path.IsPathRooted(candidate)) continue;
                if (!candidate.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                    candidate = Path.Combine(candidate, "VSTOInstaller.exe");
                if (string.Equals(Path.GetFileName(candidate), "VSTOInstaller.exe", StringComparison.OrdinalIgnoreCase)
                    && fileExists(candidate)) return candidate;
            }
            foreach (string root in commonRoots)
            {
                if (string.IsNullOrWhiteSpace(root)) continue;
                // The .NET 4 add-ins require the VS 2010 (10.0) runtime or later.
                for (int version = 14; version >= 10; version--)
                {
                    string candidate = Path.Combine(root, @"Microsoft Shared\VSTO\" + version + ".0", "VSTOInstaller.exe");
                    if (fileExists(candidate)) return candidate;
                }
            }
            return null;
        }

        internal static bool IsOwnedUrl(string value, string manifestName)
        {
            if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(manifestName)) return false;
            int suffix = value.IndexOf('|');
            if (suffix >= 0) value = value.Substring(0, suffix);
            Uri uri;
            return Uri.TryCreate(value, UriKind.Absolute, out uri) && uri.IsFile
                && string.IsNullOrEmpty(uri.Query) && string.IsNullOrEmpty(uri.Fragment)
                && string.Equals(Path.GetFileName(uri.LocalPath), manifestName, StringComparison.OrdinalIgnoreCase);
        }
    }

    internal static class SubscriptionCleanup
    {
        internal static InstallResult Remove(IEnumerable<string> subscriptionUrls, string manifestName,
            Func<string, int> uninstall, Action<string> log)
        {
            HashSet<string> processed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string value in subscriptionUrls)
            {
                if (!RuntimeLocator.IsOwnedUrl(value, manifestName)) continue;
                string raw = value;
                int suffix = raw.IndexOf('|');
                if (suffix >= 0) raw = raw.Substring(0, suffix);
                string url = new Uri(raw).AbsoluteUri;
                if (!processed.Add(url)) continue;
                InstallResult result = InstallResult.FromExitCode(uninstall(url), "注销旧订阅");
                if (result.Status != InstallStatus.Success) {
                    log("旧订阅注销未完成，停止操作并保留程序文件：" + url);
                    return result;
                }
                log("已成功注销本加载项订阅：" + url);
            }
            return new InstallResult(InstallStatus.Success, "本加载项旧订阅检查完成。", 0);
        }
    }

    internal sealed class InstallActions
    {
        internal Func<string> FindRuntime;
        internal Func<string, InstallResult> CleanupSubscriptions;
        internal Action ReleasePayload;
        internal Func<string, int> InvokeInstall;
        internal Action Register;
        internal Action Complete;
    }

    internal sealed class UninstallActions
    {
        internal Func<string> FindRuntime;
        internal Func<string, InstallResult> CleanupSubscriptions;
        internal Action RemovePayloadAndRegistration;
    }

    // The order is tested with injected fake operations. Only an official exit code
    // of zero can activate a newly released add-in.
    internal static class InstallFlow
    {
        internal static InstallResult Install(InstallActions actions)
        {
            try
            {
                string tool = actions.FindRuntime();
                if (tool == null) return new InstallResult(InstallStatus.MissingRuntime, RuntimeLocator.MissingMessage, null);
                InstallResult clean = actions.CleanupSubscriptions(tool);
                if (clean.Status != InstallStatus.Success) return clean;
                actions.ReleasePayload();
                InstallResult result = InstallResult.FromExitCode(actions.InvokeInstall(tool), "安装");
                if (result.Status != InstallStatus.Success) return result;
                actions.Register();
                actions.Complete();
                return result;
            }
            catch (Exception ex) { return new InstallResult(InstallStatus.Failed, "安装失败：" + ex.Message, null); }
        }

        internal static InstallResult Uninstall(UninstallActions actions)
        {
            try
            {
                string tool = actions.FindRuntime();
                if (tool == null) return new InstallResult(InstallStatus.MissingRuntime, RuntimeLocator.MissingMessage, null);
                InstallResult clean = actions.CleanupSubscriptions(tool);
                if (clean.Status != InstallStatus.Success) return clean;
                actions.RemovePayloadAndRegistration();
                return new InstallResult(InstallStatus.Success, "卸载完成。", 0);
            }
            catch (Exception ex) { return new InstallResult(InstallStatus.Failed, "卸载失败：" + ex.Message, null); }
        }
    }
}
