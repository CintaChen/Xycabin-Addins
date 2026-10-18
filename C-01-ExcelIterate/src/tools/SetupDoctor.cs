using System;
using Id = WpsShared.ProjectInfo;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml;
using Microsoft.Win32;

// =====================================================================
// SetupDoctor.exe —— 共用安装环境诊断（发给同事双击运行）
//
//   排查"安装后打开 WPS 不弹「自定义项安装程序」"的常见原因：
//     1. WPS 没完全退出重启（单实例多标签，后台进程不重启就不重扫加载项）
//     2. 安装器没真正跑成功（被 SmartScreen/杀软拦，或没点「安装」）
//     3. 缺 VSTO 运行时（Microsoft Visual Studio 2010 Tools for Office Runtime）
//     4. WPS 版本/渠道差异（过旧或定制版可能不支持 VSTO 加载项）
//
//   输出：控制台打印 + 桌面落盘「<程序集名>_诊断报告.txt」，
//   让同事把这个 txt 发回来即可。
// =====================================================================

namespace WpsShared.Tools
{
    internal static class SetupDoctor
    {
        private const string AddinKeyPath =
            Id.AddinKeyPath;

        private static readonly StringBuilder R = new StringBuilder();
        private static readonly StringBuilder Advice = new StringBuilder();

        [STAThread]
        private static int Main(string[] args)
        {
            try { Console.OutputEncoding = Encoding.UTF8; } catch { }

            R.AppendLine(Id.DisplayName + " —— 安装环境诊断报告");
            R.AppendLine("生成时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            R.AppendLine(new string('=', 60));

            CheckOsAndDotNet();
            CheckWpsInstall();
            CheckWpsBitness();
            CheckVstorRuntime();
            CheckAddinRegistration();
            CheckTrustAndPayload();
            CheckWpsProcess();

            R.AppendLine();
            R.AppendLine(new string('=', 60));
            R.AppendLine("【给作者看的结论建议】");
            R.AppendLine(Advice.Length > 0
                ? Advice.ToString()
                : "各项检查未见明显异常；请把本报告完整发给插件作者进一步分析。");

            string outPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Id.AssemblyName + "_诊断报告.txt");
            for (int i = 0; i + 1 < args.Length; i++)
                if (args[i] == "--output") outPath = Path.GetFullPath(args[i + 1]);
            try
            {
                File.WriteAllText(outPath, R.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                R.AppendLine("（桌面写入失败：" + ex.Message + "，改写到临时目录）");
                outPath = Path.Combine(Path.GetTempPath(), Id.AssemblyName + "_诊断报告.txt");
                try { File.WriteAllText(outPath, R.ToString(), Encoding.UTF8); } catch { }
            }

            Console.Write(R.ToString());
            Console.WriteLine();
            Console.WriteLine("报告文件已保存：" + outPath);
            Console.WriteLine("请把这个 txt 文件发给你的同事（插件作者）。");
            Console.WriteLine();
            if (Array.IndexOf(args, "--no-pause") < 0)
            {
                Console.Write("按任意键退出…");
                try { Console.ReadKey(true); } catch { }
            }
            return 0;
        }

        // --------------------------- ① 系统与 .NET ---------------------------

        private static void CheckOsAndDotNet()
        {
            Section("① 系统与 .NET Framework");
            try
            {
                R.AppendLine("操作系统：" + Environment.OSVersion
                    + (Environment.Is64BitOperatingSystem ? "（64 位）" : "（32 位）"));
            }
            catch { }
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
                {
                    if (k == null)
                    {
                        R.AppendLine(".NET Framework 4.x：未检测到 Full 配置（异常）");
                        Advice.AppendLine("· .NET Framework 4.x 缺失或未装全，安装器与加载项都无法运行；请先安装 .NET Framework 4.8。");
                        return;
                    }
                    object rel = k.GetValue("Release");
                    int v = rel == null ? 0 : Convert.ToInt32(rel);
                    R.AppendLine(".NET Framework 4.x Release = " + v + "（" + DescribeDotNet(v) + "）");
                    if (v > 0 && v < 378389)
                        Advice.AppendLine("· .NET Framework 低于 4.5，需要升级后才能运行加载项。");
                }
            }
            catch (Exception ex)
            {
                R.AppendLine(".NET 检测出错：" + ex.Message);
            }
        }

        private static string DescribeDotNet(int release)
        {
            if (release >= 533325) return "4.8.1 或更高";
            if (release >= 528040) return "4.8";
            if (release >= 461808) return "4.7.2";
            if (release >= 461308) return "4.7.1";
            if (release >= 460798) return "4.7";
            if (release >= 394802) return "4.6.2";
            if (release >= 394254) return "4.6.1";
            if (release >= 393295) return "4.6";
            if (release >= 379893) return "4.5.2";
            if (release >= 378675) return "4.5.1";
            if (release >= 378389) return "4.5";
            return "低于 4.5";
        }

        // --------------------------- ② WPS 安装与版本 ---------------------------

        private static void CheckWpsInstall()
        {
            Section("② WPS 安装与版本");
            bool found = false;
            found |= ScanUninstall(RegistryHive.CurrentUser, RegistryView.Default, "HKCU");
            found |= ScanUninstall(RegistryHive.LocalMachine, RegistryView.Registry32, "HKLM-32");
            found |= ScanUninstall(RegistryHive.LocalMachine, RegistryView.Registry64, "HKLM-64");

            // 常见安装目录下也能看出版本
            try
            {
                string local = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Kingsoft", "WPS Office");
                if (Directory.Exists(local))
                {
                    string[] dirs = Directory.GetDirectories(local);
                    R.AppendLine("WPS 本地版本目录：" + string.Join("、",
                        Array.ConvertAll(dirs, Path.GetFileName)));
                    found = true;
                }
            }
            catch { }

            if (!found)
            {
                R.AppendLine("未在常见位置找到 WPS 安装信息（可能装在非标准位置）。");
                Advice.AppendLine("· 未检测到 WPS 安装信息——请确认这台电脑装的是支持加载项的" + Id.HostDisplayName + "（且不是过旧版本）。");
            }
            else
            {
                R.AppendLine("（提示：WPS 11.x 或更早版本、以及部分政企定制版不支持 VSTO 加载项；"
                    + "作者开发验证用的是 12.1.0.28488。版本明显偏低时请升级 WPS 个人版。）");
            }
        }

        private static bool ScanUninstall(RegistryHive hive, RegistryView view, string tag)
        {
            bool found = false;
            try
            {
                using (RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, view))
                using (RegistryKey un = baseKey.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"))
                {
                    if (un == null) return false;
                    foreach (string sub in un.GetSubKeyNames())
                    {
                        try
                        {
                            using (RegistryKey k = un.OpenSubKey(sub))
                            {
                                if (k == null) continue;
                                string name = k.GetValue("DisplayName") as string;
                                if (string.IsNullOrEmpty(name)) continue;
                                if (name.IndexOf("WPS", StringComparison.OrdinalIgnoreCase) < 0)
                                    continue;
                                string ver = k.GetValue("DisplayVersion") as string;
                                R.AppendLine(tag + " 找到：" + name
                                    + (string.IsNullOrEmpty(ver) ? "" : "  版本 " + ver));
                                found = true;
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return found;
        }

        // --------------------------- ②b WPS 位数 ---------------------------

        private static void CheckWpsBitness()
        {
            Section("②b WPS 程序位数（32/64 位对 VSTO 支持可能有差异）");
            string exe = FindWpsExe();
            if (exe == null)
            {
                R.AppendLine("未在常见目录找到 wps.exe。");
                return;
            }
            R.AppendLine("wps.exe：" + exe);
            try
            {
                using (FileStream fs = new FileStream(
                    exe, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (BinaryReader br = new BinaryReader(fs))
                {
                    fs.Seek(0x3C, SeekOrigin.Begin);
                    int pe = br.ReadInt32();
                    fs.Seek(pe + 4, SeekOrigin.Begin);
                    int machine = br.ReadUInt16();
                    string desc = machine == 0x8664 ? "64 位 (x64)"
                        : (machine == 0x14C ? "32 位 (x86)" : "未知 0x" + machine.ToString("X4"));
                    R.AppendLine("位数：" + desc);
                    if (machine == 0x8664)
                        Advice.AppendLine("· 这台机器的 WPS 是【64 位】版——作者的验证环境是 32 位 WPS，"
                            + "64 位版对 VSTO 加载项的支持可能不同。建议改装 32 位 WPS 个人版再试。");
                }
            }
            catch (Exception ex)
            {
                R.AppendLine("读取 PE 头失败：" + ex.Message);
            }
        }

        private static string FindWpsExe()
        {
            string[] roots = new string[3];
            try
            {
                roots[0] = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Kingsoft", "WPS Office");
                roots[1] = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    "Kingsoft", "WPS Office");
                roots[2] = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    "Kingsoft", "WPS Office");
            }
            catch { }
            foreach (string root in roots)
            {
                try
                {
                    if (string.IsNullOrEmpty(root) || !Directory.Exists(root)) continue;
                    foreach (string ver in Directory.GetDirectories(root))
                    {
                        string p = Path.Combine(ver, "office6", "wps.exe");
                        if (File.Exists(p)) return p;
                    }
                    string direct = Path.Combine(root, "office6", "wps.exe");
                    if (File.Exists(direct)) return direct;
                }
                catch { }
            }
            return null;
        }

        // --------------------------- ③ VSTO 运行时 ---------------------------

        private static void CheckVstorRuntime()
        {
            Section("③ VSTO 运行时（加载 .vsto 所必需）");
            string installer = RuntimeLocator.FindInstalledTool();
            R.AppendLine("VSTOInstaller（两种注册表视图的 InstallerPath + 默认目录）："
                + (installer ?? "未找到"));
            if (installer == null)
                Advice.AppendLine("· 未找到官方 VSTOInstaller。先从微软官方渠道安装或修复 VSTO Runtime，"
                    + "再运行加载项安装器；直接写入信任记录不能代替运行库。"
                    + " https://learn.microsoft.com/en-us/visualstudio/vsto/how-to-install-the-visual-studio-tools-for-office-runtime-redistributable");
            string gac = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "Microsoft.NET", "assembly", "GAC_MSIL");
            string[] need = new string[]
            {
                "Microsoft.VisualStudio.Tools.Applications.Runtime",
                "Microsoft.Office.Tools",
                "Microsoft.Office.Tools.Common",
                "Microsoft.Office.Tools.v4.0.Framework",
                "Microsoft.Office.Tools." + Id.Host,
            };
            int missing = 0;
            foreach (string name in need)
            {
                bool ok = Directory.Exists(Path.Combine(gac, name));
                R.AppendLine((ok ? "  [有] " : "  [缺] ") + name);
                if (!ok) missing++;
            }
            if (missing > 0)
            {
                R.AppendLine("缺少 " + missing + " 个运行时程序集。WPS 加载 .vsto 前会做先决条件检查，"
                    + "缺失时可能直接不弹安装提示。");
                Advice.AppendLine("· 缺 VSTO 运行时：下载安装微软的 vstor_redist（Visual Studio 2010 Tools "
                    + "for Office Runtime），或安装过 Microsoft Office 的机器自带。装完后重开 WPS。");
            }
            else
            {
                R.AppendLine(installer == null ? "程序集存在，但官方安装工具缺失，运行库可能不完整。"
                    : "已找到运行时程序集和官方安装工具；真实启用仍需宿主验收。");
            }
        }

        // --------------------------- ④ 加载项注册表 ---------------------------

        private static void CheckAddinRegistration()
        {
            Section("④ 加载项注册表（安装器写入的挂载点）");

            // 目录记忆只说明曾记录过位置，旧安装器会在官方授权完成前写入。
            // 它不证明安装成功，也不能据此断定注册表被清理。
            string lastDir = null;
            string lastTime = null;
            try
            {
                using (RegistryKey sk = Registry.CurrentUser.OpenSubKey(
                    Id.SetupKeyPath))
                {
                    if (sk != null)
                    {
                        lastDir = sk.GetValue("LastInstallDir") as string;
                        lastTime = sk.GetValue("LastInstalledAt") as string;
                    }
                }
            }
            catch { }

            RegistryKey k = null;
            try { k = Registry.CurrentUser.OpenSubKey(AddinKeyPath); }
            catch { }
            if (k == null)
            {
                R.AppendLine("HKCU\\" + AddinKeyPath + "：不存在。");
                if (!string.IsNullOrEmpty(lastDir))
                {
                    R.AppendLine("但安装器记忆键存在：上次安装于 " + (lastTime ?? "?")
                        + " → " + lastDir);
                    R.AppendLine("→ 有目录记忆但没有挂载点，不能仅凭该记忆确认此前已安装成功。");
                    Advice.AppendLine("· 先核对官方 VSTO 安装结果和安装日志；补齐运行库后重新运行安装器。");
                }
                else
                {
                    // 再看看默认安装目录里有没有释放出来的文件
                    string defDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "Programs", Id.AddinName);
                    int files = 0;
                    try
                    {
                        if (Directory.Exists(defDir))
                            files = Directory.GetFiles(defDir).Length;
                    }
                    catch { }
                    if (files > 0)
                    {
                        R.AppendLine("默认目录 " + defDir + " 里有 " + files
                            + " 个文件，但两个注册表键都不存在。");
                        R.AppendLine("→ 文件释放了、注册表没写上：写入那一步被拦截。");
                        Advice.AppendLine("· 注册表写入被拦：先暂时退出 360/电脑管家等安全软件，再重新运行安装器点「安装」，"
                            + "装完再开回来。");
                    }
                    else
                    {
                        R.AppendLine("安装器记忆键也不存在、默认目录里也没有文件。");
                        R.AppendLine("→ 若安装时选的是【自定义目录】，默认目录无文件属正常；此时最常见的");
                        R.AppendLine("  原因是：用【以管理员身份运行】打开过安装器——文件照常释放到了所选");
                        R.AppendLine("  目录，但注册表写进了【管理员账户】，当前登录账户的 WPS 看不到。");
                        Advice.AppendLine("· 挂载点缺失的最可能原因 = 以管理员身份运行导致写错账户：请【普通双击】安装器"
                            + "（不要右键以管理员身份运行），点【安装】（目录保持原选择即可），"
                            + "确认日志显示「安装完成！下一步…」，再完全退出 WPS 重开。"
                            + "若仍不行，检查 360/电脑管家是否拦截了注册表写入。");
                    }
                }
                return;
            }
            using (k)
            {
                string manifest = k.GetValue("Manifest") as string;
                object lb = k.GetValue("LoadBehavior");
                R.AppendLine("Manifest     = " + (manifest ?? "(空)"));
                R.AppendLine("LoadBehavior = " + (lb == null ? "(空)" : lb.ToString())
                    + DescribeLoadBehavior(lb));
                R.AppendLine("FriendlyName = " + (k.GetValue("FriendlyName") as string ?? "(空)"));

                string dir = DirFromManifest(manifest);
                if (dir != null)
                {
                    R.AppendLine("安装目录：" + dir);
                    string[] payload = new string[]
                    {
                        Id.AssemblyName + ".dll",
                        "Microsoft.Office.Tools.Common.v4.0.Utilities.dll",
                        "Interop.Microsoft.Office.Core.dll",
                        Id.AssemblyName + ".dll.manifest",
                        Id.PayloadVstoName,
                    };
                    int missing = 0;
                    foreach (string name in payload)
                    {
                        string p = Path.Combine(dir, name);
                        bool ok = File.Exists(p);
                        R.AppendLine((ok ? "  [有] " : "  [缺] ") + p);
                        if (!ok) missing++;
                    }
                    if (missing > 0)
                        Advice.AppendLine("· 安装目录里缺 " + missing + " 个文件（可能被清理软件误删）：重新运行安装器再装一次。");
                }
                else
                {
                    Advice.AppendLine("· Manifest 值格式异常，无法解析安装目录：重新运行安装器再装一次。");
                }

                int lbv = lb == null ? -1 : Convert.ToInt32(lb);
                if (lbv == 2)
                    Advice.AppendLine("· LoadBehavior=2：WPS 曾尝试加载但失败，已把加载项禁用。常见原因是缺 VSTO 运行时"
                        + "或清单校验失败——先按 ③ 补齐运行时，再重启 WPS 观察弹出的错误框内容（很关键）。");
                else if (lbv == 0)
                    Advice.AppendLine("· LoadBehavior=0：加载项被禁用。重装一次并把该值改回 3。");
            }
        }

        private static string DescribeLoadBehavior(object lb)
        {
            if (lb == null) return "";
            int v = Convert.ToInt32(lb);
            if (v == 3) return "（正常：启动时加载）";
            if (v == 2) return "（注意：已被禁用，上次加载失败过）";
            if (v == 0) return "（注意：完全禁用）";
            return "（异常值）";
        }

        private static string DirFromManifest(string manifest)
        {
            try
            {
                if (string.IsNullOrEmpty(manifest)) return null;
                string v = manifest;
                int bar = v.IndexOf('|');
                if (bar >= 0) v = v.Substring(0, bar);
                Uri uri;
                return Uri.TryCreate(v, UriKind.Absolute, out uri) && uri.IsFile
                    && string.IsNullOrEmpty(uri.Query) && string.IsNullOrEmpty(uri.Fragment)
                    ? Path.GetDirectoryName(uri.LocalPath) : null;
            }
            catch { return null; }
        }

        // --------------------------- ⑤ WPS 进程 ---------------------------

        private static void CheckWpsProcess()
        {
            Section("⑤ WPS 进程状态");
            string[] names = new string[] { "wps", "et", "wpp" };
            int running = 0;
            foreach (string n in names)
            {
                try
                {
                    Process[] ps = Process.GetProcessesByName(n);
                    if (ps.Length > 0)
                    {
                        R.AppendLine("  正在运行：" + n + ".exe × " + ps.Length);
                        running += ps.Length;
                    }
                }
                catch { }
            }
            if (running > 0)
            {
                R.AppendLine("WPS 是单实例程序：只要还有 wps.exe 在跑（包括托盘/后台），"
                    + "新开的窗口不会重新扫描加载项注册表，自然不弹安装提示。");
                Advice.AppendLine("· 有 WPS 进程在后台运行：任务栏右下角退出 WPS，或在任务管理器结束全部 wps.exe 后，"
                    + "再重新打开" + Id.HostDisplayName + "。");
            }
            else
            {
                R.AppendLine("当前没有 WPS 进程在运行（好——下次启动会完整重扫加载项）。");
            }
        }

        // --------------------------- 输出 ---------------------------

        // --------------------------- ⑥ 四要素自检 ---------------------------
        // 决定加载项能否被 WPS 加载的就是这四件事，缺一不可：
        //   ① 5 个载荷确实躺在安装目录里
        //   ② 注册表 Manifest 与 LoadBehavior=3
        //   ③ ★ WPS 白名单 AddinsWL 里有本加载项名字（最容易被漏，且症状最隐蔽）
        //   ④ VSTO 侧认可本加载项（信任包含列表条目 / ClickOnce 订阅）
        // 顺带把签名算法也报出来（SHA-1 会被较新的 VSTO 判为不可信）。

        private static void CheckTrustAndPayload()
        {
            Section("⑥ 四要素自检（决定能否加载，缺一不可）");

            string manifest = null;
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(AddinKeyPath))
                {
                    if (k != null)
                    {
                        object v = k.GetValue("Manifest");
                        if (v != null) manifest = v.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                R.AppendLine("  读取加载项注册表失败：" + ex.Message);
            }

            string dir = DirFromManifest(manifest);
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
            {
                R.AppendLine("  未能定位安装目录（可能尚未安装），四要素自检跳过。");
                return;
            }
            R.AppendLine("  安装目录：" + dir);

            // ① 载荷
            string[] pay = new string[]
            {
                Id.AssemblyName + ".dll",
                "Microsoft.Office.Tools.Common.v4.0.Utilities.dll",
                "Interop.Microsoft.Office.Core.dll",
                Id.AssemblyName + ".dll.manifest",
                Id.PayloadVstoName,
            };
            bool filesOk = true;
            foreach (string f in pay)
            {
                string p = Path.Combine(dir, f);
                bool e = File.Exists(p);
                if (!e) filesOk = false;
                R.AppendLine("   ①载荷 " + (e ? "OK   " : "缺失!") + " " + f
                    + (e ? "  (" + new FileInfo(p).Length + " 字节)" : ""));
            }

            // ② |vstolocal（信息项）+ LoadBehavior（硬性项）
            bool vstoLocal = !string.IsNullOrEmpty(manifest)
                && manifest.EndsWith("|vstolocal", StringComparison.OrdinalIgnoreCase);
            bool lbOk = false;
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(AddinKeyPath))
                {
                    if (k != null)
                    {
                        object lbv = k.GetValue("LoadBehavior");
                        lbOk = (lbv != null && lbv.ToString() == "3");
                    }
                }
            }
            catch { }
            R.AppendLine("   ②Manifest 以 |vstolocal 结尾 : " + vstoLocal
                + "（标准弹窗安装后 VSTO 可能自己改成不带后缀，不影响加载）");
            R.AppendLine("     LoadBehavior=3（启动加载）: " + lbOk);
            if (!lbOk)
                R.AppendLine("      !! LoadBehavior 不为 3 —— 到 ④ 节看是否被降级为 2。");

            // ③ ★ WPS 白名单（最容易被漏掉的一条，务必排在其他检查之前判断）
            bool wlOk = false;
            try
            {
                using (RegistryKey root = Registry.CurrentUser.OpenSubKey(
                    Id.AddinsWlKeyPath))
                {
                    if (root == null)
                    {
                        R.AppendLine("   ③WPS 加载项白名单 AddinsWL : 键不存在");
                    }
                    else
                    {
                        object v = root.GetValue(Id.AddinName);
                        wlOk = (v != null);
                        int total = root.GetValueNames().Length;
                        R.AppendLine("   ③WPS 加载项白名单 AddinsWL : "
                            + (wlOk ? "已登记" : "★未登记 —— WPS 会整条跳过本加载项")
                            + "（白名单里共 " + total + " 个名字）");
                        if (!wlOk)
                            R.AppendLine("      症状：不报错、事件日志无记录、LoadBehavior 也不降级，"
                                + "看上去像没装。补一个空字符串值 " + Id.AddinName + " 即可。");
                    }
                }
            }
            catch (Exception ex)
            {
                R.AppendLine("   ③读取 AddinsWL 失败：" + ex.Message);
            }

            // ④ 信任包含列表
            string vstoPath = Path.Combine(dir, Id.PayloadVstoName);
            string wantUrl = new Uri(Path.GetFullPath(vstoPath)).AbsoluteUri;
            string myPub = null;
            try
            {
                myPub = PublicKeyOfVsto(vstoPath);
            }
            catch (Exception ex)
            {
                R.AppendLine("   （读取部署清单里的签名证书失败：" + ex.Message + "）");
            }

            bool trustOk = false;
            try
            {
                using (RegistryKey root = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\VSTO\Security\Inclusion"))
                {
                    if (root == null)
                    {
                        R.AppendLine("   ④VSTO 信任包含列表 : 键不存在（未登记）");
                    }
                    else
                    {
                        int n = root.GetSubKeyNames().Length;
                        foreach (string name in root.GetSubKeyNames())
                        {
                            using (RegistryKey k = root.OpenSubKey(name))
                            {
                                if (k == null) continue;
                                object u = k.GetValue("Url");
                                object pk = k.GetValue("PublicKey");
                                if (u == null) continue;
                                if (!SameFileUrl(u.ToString(), wantUrl)) continue;
                                bool same = (pk != null && myPub != null &&
                                    string.Equals(pk.ToString(), myPub, StringComparison.Ordinal));
                                trustOk = trustOk || same;
                                R.AppendLine("   ④VSTO 信任包含列表 : 已登记（" + name + "）");
                                R.AppendLine("      公钥与当前清单一致 : " + same);
                                if (!same)
                                    R.AppendLine("      !! 条目公钥与当前清单不符——多半是换过签名证书后没重装，"
                                        + "用最新安装器重装一次即可。");
                            }
                        }
                        R.AppendLine("      （包含列表里共 " + n + " 个条目，含其他插件）");
                        if (!trustOk)
                            R.AppendLine("   ④VSTO 信任包含列表 : 未找到与当前公钥匹配的条目；"
                                + "还需核对官方安装订阅和授权结果，不能仅凭此项判断会否弹窗。");
                    }
                }
            }
            catch (Exception ex)
            {
                R.AppendLine("   ④读取包含列表失败：" + ex.Message);
            }

            // ⑤ 签名算法
            try
            {
                string t = File.ReadAllText(vstoPath);
                bool sha256 = t.IndexOf("rsa-sha256", StringComparison.OrdinalIgnoreCase) >= 0;
                bool sha1 = t.IndexOf("rsa-sha1", StringComparison.OrdinalIgnoreCase) >= 0;
                int sigs = 0, i = 0;
                while ((i = t.IndexOf("<Signature", i, StringComparison.Ordinal)) >= 0)
                {
                    sigs++;
                    i += 10;
                }
                R.AppendLine("   ⑤部署清单签名 : " + sigs + " 段，算法="
                    + (sha256 ? "SHA-256（正常）"
                              : (sha1 ? "SHA-1（已废弃，可能被拒）" : "未知")));
            }
            catch { }

            bool subscription = false;
            using (RegistryKey solutions = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VSTA\Solutions"))
            {
                if (solutions != null)
                    foreach (string name in solutions.GetSubKeyNames())
                        using (RegistryKey key = solutions.OpenSubKey(name))
                            if (key != null && SameFileUrl(key.GetValue("Url") as string, wantUrl)) subscription = true;
            }
            R.AppendLine("   官方 VSTO 订阅的 URL 与当前部署清单匹配 : " + subscription
                + "（登记存在不等于真实宿主启用成功）");
            bool runtime = RuntimeLocator.FindInstalledTool() != null;
            bool all = filesOk && lbOk && wlOk && runtime && (trustOk || subscription);
            R.AppendLine();
            R.AppendLine("  结论：" + (all
                ? "发现载荷、挂载点、白名单、官方安装工具及授权登记线索；请在真实宿主中验收。"
                : "发现缺失项或无法确认授权登记，见上；当前检查不代表加载成功。"));

            if (all)
            {
                Advice.AppendLine("· 已找到本项目登记线索；仍需核对安装日志中的官方返回码及宿主真实启用结果。"
                    + "若 WPS 不显示选项卡，保存文档并完全退出后重开，再检查清单校验与运行库完整性。");
            }
            else
            {
                Advice.AppendLine("· 先补齐微软 VSTO Runtime，再运行最新的 " + Id.AssemblyName
                    + "_Setup.exe，在官方安装窗口完成授权并确认安装日志返回 0；安装器不再直接代写信任。");
            }
        }

        private static bool SameFileUrl(string left, string right)
        {
            try
            {
                Uri a, b;
                return Uri.TryCreate(left, UriKind.Absolute, out a) && a.IsFile
                    && string.IsNullOrEmpty(a.Query) && string.IsNullOrEmpty(a.Fragment)
                    && Uri.TryCreate(right, UriKind.Absolute, out b) && b.IsFile
                    && string.IsNullOrEmpty(b.Query) && string.IsNullOrEmpty(b.Fragment)
                    && string.Equals(Path.GetFullPath(a.LocalPath), Path.GetFullPath(b.LocalPath),
                        StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        /// <summary>从部署清单的 XML 签名里取证书公钥（只要公钥，不需要私钥）。</summary>
        private static string PublicKeyOfVsto(string vstoPath)
        {
            XmlDocument doc = new XmlDocument();
            doc.PreserveWhitespace = true;
            doc.Load(vstoPath);
            XmlNodeList nodes = doc.GetElementsByTagName(
                "X509Certificate", "http://www.w3.org/2000/09/xmldsig#");
            if (nodes.Count == 0) return null;
            X509Certificate2 cert =
                new X509Certificate2(Convert.FromBase64String(nodes[0].InnerText.Trim()));
            RSA rsa = cert.GetRSAPublicKey();
            if (rsa == null) return null;
            RSAParameters p = rsa.ExportParameters(false);
            return "<RSAKeyValue><Modulus>" + Convert.ToBase64String(p.Modulus)
                + "</Modulus><Exponent>" + Convert.ToBase64String(p.Exponent)
                + "</Exponent></RSAKeyValue>";
        }

        private static void Section(string title)
        {
            R.AppendLine();
            R.AppendLine(title);
            R.AppendLine(new string('-', 40));
        }
    }
}
