using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    // One controller per add-in assembly. Never retain a borrowed host application.
    internal sealed class AgentExtensionController
    {
        internal const string TabId = "tabAgentEvo";
        internal const string ToggleId = "btnAgentExtension";
        private readonly List<object> ribbons = new List<object>();
        private readonly Action<bool> save;
        private readonly Action<string> log;
        internal bool Visible { get; private set; }
        internal int RibbonCount { get { return ribbons.Count; } }

        internal AgentExtensionController(Func<bool> read, Action<bool> save, Action<string> log)
        {
            this.save = save; this.log = log;
            try { Visible = read(); }
            catch (Exception ex) { Visible = false; log("Agent 状态读取失败，使用默认隐藏：" + ex.Message); }
        }
        internal void Register(object ribbon)
        {
            if (ribbon == null) throw new ArgumentNullException("ribbon");
            if (!ribbons.Contains(ribbon)) ribbons.Add(ribbon);
        }
        private static void Refresh(object ribbon)
        {
            try
            {
                var typed = ribbon as Microsoft.Office.Core.IRibbonUI;
                if (typed != null) { typed.InvalidateControl(TabId); typed.InvalidateControl(ToggleId); }
                else { dynamic r = ribbon; r.InvalidateControl(TabId); r.InvalidateControl(ToggleId); }
            }
            catch
            {
                var typed = ribbon as Microsoft.Office.Core.IRibbonUI;
                if (typed != null) typed.Invalidate();
                else { dynamic r = ribbon; r.Invalidate(); }
            }
        }
        // Empty means success; warnings describe the resulting state precisely.
        internal string SetVisible(bool pressed)
        {
            bool previous = Visible;
            Visible = pressed;
            try
            {
                if (ribbons.Count == 0) throw new InvalidOperationException("尚未取得功能区接口");
                foreach (object ribbon in ribbons) Refresh(ribbon);
            }
            catch (Exception ex)
            {
                Visible = previous;
                foreach (object ribbon in ribbons)
                    try { Refresh(ribbon); } catch (Exception rollback) { log("Agent 回退刷新失败：" + rollback.Message); }
                string error = "无法刷新 Agent-Evo，已恢复原开关状态。\r\n" + ex.Message;
                log(error); return error;
            }
            try { save(Visible); }
            catch (Exception ex)
            {
                string warning = "本次开关已生效，但未能记住选择；下次启动可能恢复原状态。\r\n" + ex.Message;
                log(warning); return warning;
            }
            log("Agent-Evo：" + (Visible ? "显示" : "隐藏")); return "";
        }
        internal void Clear() { ribbons.Clear(); }
    }

    internal static class AgentSettings
    {
        internal static string PathForUser
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                typeof(Connect).Assembly.GetName().Name, "AgentExtension.state"); }
        }
        internal static bool Read(string path)
        {
            if (!File.Exists(path)) return false;
            bool value;
            if (!bool.TryParse(File.ReadAllText(path, Encoding.UTF8).Trim(), out value))
                throw new InvalidDataException("开关状态文件格式不正确");
            return value;
        }
        internal static void Save(string path, bool value)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temporary, value ? "true" : "false", new UTF8Encoding(false));
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }

    public partial class Connect
    {
        public Connect() { DialogSizing.Initialize(); }
        private static AgentExtensionController agentController;
        internal static Action<string> AgentWebsiteLauncher = delegate(string url) { Process.Start(url); };
        private static AgentExtensionController AgentController
        {
            get
            {
                if (agentController == null) agentController = new AgentExtensionController(
                    delegate { return AgentSettings.Read(AgentSettings.PathForUser); },
                    delegate(bool value) { AgentSettings.Save(AgentSettings.PathForUser, value); }, Log);
                return agentController;
            }
        }
        public void OnAgentRibbonLoad(object ribbonUI)
        {
            try { FormatSession.Initialize(); AgentController.Register(ribbonUI); RibbonCustomization.Register(ribbonUI); }
            catch (Exception ex) { Log("Agent 功能区加载失败：" + ex); }
        }
        public bool GetAgentTabVisible(object control) { return AgentController.Visible; }
        public bool GetAgentExtensionPressed(object control) { return AgentController.Visible; }
        public void OnToggleAgentExtension(object control, bool pressed)
        {
            string warning = AgentController.SetVisible(pressed);
            if (warning.Length != 0) UiDialog.Show(warning, "Agent扩展", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        internal static void CloseAgentExtension()
        {
            AgentDocumentWindows.CloseAll();
            CustomizationWindows.CloseAll();RibbonCustomization.Clear();
#if !WPS_WORD
            CloseFileNames();
#endif
            DialogSizesForm.CloseWindow();
            if (agentController != null) agentController.Clear();
            agentController = null;
        }
        public void OnOpenAgentDialogSizes(object control)
        {
            try {
if(!EnsureApplication())throw new InvalidOperationException("请先打开对应 WPS 文件。");
                DialogSizesForm.Open(); }
            catch (Exception ex) { Log("弹窗设置打开失败："+ex); UiDialog.Show("无法打开弹窗设置："+ex.Message,"自定配置",MessageBoxButtons.OK,MessageBoxIcon.Warning); }
        }
        public void OnOpenAgentReadme(object control) { OpenAgentDocument(false); }
        public void OnOpenAgentUserGuide(object control) { OpenAgentDocument(true); }
        private static void OpenAgentDocument(bool userGuide)
        {
            try { AgentDocumentWindows.Open(userGuide); }
            catch (Exception ex) { Log("Agent 文档读取失败：" + ex); UiDialog.Show("无法读取内嵌文档：" + ex.Message, "开发流程", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        public void OnAgentHome(object control)
        {
            string path = "";
            try
            {
                path = AgentLocalFiles.ExportHome();
                AgentLocalFiles.BrowserLauncher(path);
                Log("XyCabin：" + path);
            }
            catch (Exception ex)
            {
                Log("XyCabin打开失败：" + ex);
                UiDialog.Show("无法打开 XyCabin：" + ex.Message + (path.Length == 0 ? "" : "\r\n文件路径：" + path),
                    "XyCabin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private static void OpenAgentUrl(string id)
        {
            try
            {
                string url = RibbonConfig.GetUrl(id); Uri uri;
                if (!Uri.TryCreate(url, UriKind.Absolute, out uri) || (uri.Scheme != "https" && uri.Scheme != "http"))
                    throw new InvalidDataException("网站入口必须使用 http 或 https 地址");
                AgentWebsiteLauncher(url); Log("调用 API：" + url);
            }
            catch (Exception ex) { Log("调用 API 打开失败：" + ex); UiDialog.Show("无法打开网站：" + ex.Message, "调用API", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        private static IPictureDisp GetAgentImage(string id)
        {
            IPictureDisp custom = RibbonConfig.GetImage(id);
            return custom ?? MakeRibbonIcon(RibbonConfig.BuiltinIconKind(id) ?? "agent", Ui.Blue, RibbonConfig.GetImageSize(id));
        }
        public IPictureDisp GetImage_btnAgentExtension(object control) { return GetAgentImage("btnAgentExtension"); }
        public IPictureDisp GetImage_menuAgentApi(object control) { return GetAgentImage("menuAgentApi"); }
        public IPictureDisp GetImage_menuAgentCloud(object control) { return GetAgentImage("menuAgentCloud"); }
        public IPictureDisp GetImage_menuAgentSites(object control) { return GetAgentImage("menuAgentSites"); }
        public IPictureDisp GetImage_menuAgentChina(object control) { return GetAgentImage("menuAgentChina"); }
        public IPictureDisp GetImage_menuAgentOverseas(object control) { return GetAgentImage("menuAgentOverseas"); }
        public IPictureDisp GetImage_menuAgentCustomization(object control) { return GetAgentImage("menuAgentCustomization"); }
        public IPictureDisp GetImage_btnAgentDialogSizes(object control) { return GetAgentImage("btnAgentDialogSizes"); }
        public IPictureDisp GetImage_menuAgentWorkflow(object control) { return GetAgentImage("menuAgentWorkflow"); }
        public IPictureDisp GetImage_btnAgentReadme(object control) { return GetAgentImage("btnAgentReadme"); }
        public IPictureDisp GetImage_btnAgentUserGuide(object control) { return GetAgentImage("btnAgentUserGuide"); }
        public IPictureDisp GetImage_btnAgentHome(object control) { return GetAgentImage("btnAgentHome"); }
        public IPictureDisp GetImage_btnAgentChatGPT(object control) { return GetAgentImage("btnAgentChatGPT"); }
        public IPictureDisp GetImage_btnAgentClaude(object control) { return GetAgentImage("btnAgentClaude"); }
        public IPictureDisp GetImage_btnAgentWorkBuddy(object control) { return GetAgentImage("btnAgentWorkBuddy"); }
        public IPictureDisp GetImage_btnAgentTrae(object control) { return GetAgentImage("btnAgentTrae"); }
        public IPictureDisp GetImage_btnAgentQoder(object control) { return GetAgentImage("btnAgentQoder"); }
        public IPictureDisp GetImage_btnAgentGemini(object control) { return GetAgentImage("btnAgentGemini"); }
        public IPictureDisp GetImage_btnAgentDeepSeek(object control) { return GetAgentImage("btnAgentDeepSeek"); }
        public IPictureDisp GetImage_btnAgentDoubao(object control) { return GetAgentImage("btnAgentDoubao"); }
        public IPictureDisp GetImage_btnAgentKimi(object control) { return GetAgentImage("btnAgentKimi"); }
        public IPictureDisp GetImage_btnAgentBaidu(object control) { return GetAgentImage("btnAgentBaidu"); }
        public IPictureDisp GetImage_btnAgentAliyun(object control) { return GetAgentImage("btnAgentAliyun"); }
        public void OnOpenUrl_btnAgentChatGPT(object control) { OpenAgentUrl("btnAgentChatGPT"); }
        public void OnOpenUrl_btnAgentClaude(object control) { OpenAgentUrl("btnAgentClaude"); }
        public void OnOpenUrl_btnAgentWorkBuddy(object control) { OpenAgentUrl("btnAgentWorkBuddy"); }
        public void OnOpenUrl_btnAgentTrae(object control) { OpenAgentUrl("btnAgentTrae"); }
        public void OnOpenUrl_btnAgentQoder(object control) { OpenAgentUrl("btnAgentQoder"); }
        public void OnOpenUrl_btnAgentGemini(object control) { OpenAgentUrl("btnAgentGemini"); }
        public void OnOpenUrl_btnAgentDeepSeek(object control) { OpenAgentUrl("btnAgentDeepSeek"); }
        public void OnOpenUrl_btnAgentDoubao(object control) { OpenAgentUrl("btnAgentDoubao"); }
        public void OnOpenUrl_btnAgentKimi(object control) { OpenAgentUrl("btnAgentKimi"); }
        public void OnOpenUrl_btnAgentBaidu(object control) { OpenAgentUrl("btnAgentBaidu"); }
        public void OnOpenUrl_btnAgentAliyun(object control) { OpenAgentUrl("btnAgentAliyun"); }
    }
}
