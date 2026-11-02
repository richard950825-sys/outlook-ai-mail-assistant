using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using OutlookAIAddin.Interop;

namespace OutlookAIAddin
{
    /// <summary>
    /// Outlook COM 加载项入口：
    ///  - IDTExtensibility2：Outlook 连接/断开
    ///  - IRibbonExtensibility：功能区自定义（选项卡 + 邮件右键菜单）
    /// Ribbon 回调统一用 object 接收控件参数，避免 Office 传参时接口封送问题。
    /// </summary>
    [ComVisible(true)]
    [Guid("8F5D2E1A-4C6B-4A87-9E3D-1B2C3D4E5F60")]
    [ProgId("OutlookAIAddin.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private const string RibbonResource = "OutlookAIAddin.resources.Ribbon.xml";
        private const string RibbonWordResource = "OutlookAIAddin.resources.RibbonWord.xml";
        internal static bool IsOutlookHost;
        internal static object WordAppObject;
        private const string FallbackRibbon =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs/></ribbon></customUI>";

        internal static object AppObject;
        private static IRibbonUI _ribbonUi;
        private static string _ribbonXml;
        // 诊断开关：OAI_PROBE=1 只加载 IDT；=2 加载完整功能区但返回空图标
        private static readonly string ProbeMode = Environment.GetEnvironmentVariable("OAI_PROBE") ?? "";

        public Connect()
        {
            Log.Info("Connect 实例已创建");
        }

        /* ---------------- IDTExtensibility2 ---------------- */

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                string hostName = "";
                try { dynamic hostApp = Application; hostName = Convert.ToString(hostApp.Name) ?? ""; } catch { }
                if (hostName.IndexOf("Word", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    WordAppObject = Application;
                    Log.Info("Word 宿主 OnConnection（可能是 Outlook 内嵌 Word 引擎）: " + hostName);
                    return;
                }
                IsOutlookHost = true;
                AppObject = Application;
                Log.Info("OnConnection: " + ConnectMode + ", probe=" + ProbeMode);
                if (ProbeMode == "1") { return; }

                try { System.Windows.Forms.Application.EnableVisualStyles(); }
                catch { /* 宿主可能已经初始化过 */ }

                UiHost.Reset();
            }
            catch (Exception ex)
            {
                Log.Error("OnConnection", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                Log.Info("OnDisconnection: " + RemoveMode);
                UiHost.CloseAll();
            }
            catch (Exception ex)
            {
                Log.Error("OnDisconnection", ex);
            }
            finally
            {
                AppObject = null;
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            Log.Info("OnAddInsUpdate");
        }

        public void OnStartupComplete(ref Array custom)
        {
            Log.Info("OnStartupComplete");
            if (!IsOutlookHost) { return; }
            Log.Info("正文菜单使用原生 Ribbon：ContextMenuReadOnlyMailText；不安装窗口钩子");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            try
            {
                Log.Info("OnBeginShutdown");
                UiHost.CloseAll();
            }
            catch (Exception ex)
            {
                Log.Error("OnBeginShutdown", ex);
            }
        }

        /* ---------------- IRibbonExtensibility ---------------- */

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                Log.Info("GetCustomUI: " + RibbonID);
                if (RibbonID != null && RibbonID.StartsWith("Microsoft.Word", StringComparison.OrdinalIgnoreCase))
                {
                    Log.Info("返回 Word 版功能区（contextMenus）");
                    return LoadResourceText(RibbonWordResource);
                }
                if (_ribbonXml == null)
                {
                    _ribbonXml = LoadResourceText(RibbonResource);
                }
                if (ProbeMode == "1")
                {
                    return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                         + "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs/></ribbon></customUI>";
                }
                return _ribbonXml;
            }
            catch (Exception ex)
            {
                Log.Error("GetCustomUI", ex);
                return FallbackRibbon;
            }
        }

        /* ---------------- 原生右键菜单回调（阅读窗口 / 阅读窗格） ---------------- */

        public void OnNativeReadSummary(object control)
        {
            RunSafely("OnNativeReadSummary", delegate { OpenReadSelectionAction("summary", control); });
        }

        public void OnNativeReadTranslate(object control)
        {
            RunSafely("OnNativeReadTranslate", delegate { OpenReadSelectionAction("translate", control); });
        }

        private static void OpenReadSelectionAction(string intent, object control)
        {
            object item;
            string selection = WordContextMenuHook.CaptureReadSelection(SafeContext(control), out item);
            Log.Info("原生右键菜单回调(阅读): intent=" + intent + ", control=" + SafeControlId(control)
                     + ", selectionLength=" + selection.Length + ", hasItem=" + (item != null));
            // 菜单明确标注选中文字：没有选区时不能悄悄退回整封邮件。
            if (item == null || string.IsNullOrWhiteSpace(selection))
            {
                System.Windows.Forms.MessageBox.Show("请先在邮件正文中选中文字，再使用 AI 摘要或 AI 翻译。",
                    "AI 邮件助手", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }
            UiHost.OpenAssistantWithSelection(intent, selection, item);
        }

        /* ---------------- 原生右键菜单回调（撰写窗口） ---------------- */

        public void OnNativePolishZh(object control)
        {
            RunSafely("OnNativePolishZh", delegate { OpenSelectionAction("polish_zh"); });
        }

        public void OnNativePolishEn(object control)
        {
            RunSafely("OnNativePolishEn", delegate { OpenSelectionAction("polish_en"); });
        }

        private static void OpenSelectionAction(string intent)
        {
            string selection = WordContextMenuHook.CaptureComposeSelection();
            Log.Info("原生右键菜单回调: intent=" + intent + ", selectionLength="
                     + (selection == null ? 0 : selection.Length));
            UiHost.OpenAssistantWithSelection(intent, selection ?? "");
        }

        /* ---------------- Word 原生右键菜单回调 ---------------- */

        public void OnWordPolishZh(object control)
        {
            Log.Info("Word 右键菜单回调: 润色成中文");
        }

        public void OnWordPolishEn(object control)
        {
            Log.Info("Word 右键菜单回调: 润色成英文");
        }

        /* ---------------- Ribbon 回调（参数用 object，避免接口封送问题） ---------------- */

        public void OnRibbonLoad(object ribbonUi)
        {
            try
            {
                if (ribbonUi == null) { return; }
                IRibbonUI ui = ribbonUi as IRibbonUI;
                if (ui == null)
                {
                    try { ui = (IRibbonUI)ribbonUi; } catch { ui = null; }
                }
                _ribbonUi = ui;
                Log.Info("Ribbon onLoad 完成: " + (_ribbonUi != null ? "ok" : "null"));
            }
            catch (Exception ex)
            {
                Log.Error("OnRibbonLoad", ex);
            }
        }

        public object GetImage(object control)
        {
            try
            {
                if (ProbeMode == "2") { return null; }
                string id = SafeControlId(control);
                bool small = !string.IsNullOrEmpty(id)
                             && id.StartsWith("ctx", StringComparison.OrdinalIgnoreCase);
                return PictureConverter.GetImage(small
                    ? "OutlookAIAddin.resources.icon-16.png"
                    : "OutlookAIAddin.resources.icon-32.png");
            }
            catch (Exception ex)
            {
                Log.Warn("GetImage 失败：" + ex.Message);
                return null;
            }
        }

        public void OnOpenAssistant(object control)
        {
            RunSafely("OnOpenAssistant", delegate { UiHost.OpenAssistant("open", null); });
        }

        public void OnSummary(object control)
        {
            object item = SafeContext(control);
            RunSafely("OnSummary", delegate { UiHost.OpenAssistant("summary", item); });
        }

        public void OnTranslate(object control)
        {
            object item = SafeContext(control);
            RunSafely("OnTranslate", delegate { UiHost.OpenAssistant("translate", item); });
        }

        public void OnPolish(object control)
        {
            RunSafely("OnPolish", delegate { UiHost.OpenAssistant("polish", null); });
        }

        public void OnWriteMail(object control)
        {
            RunSafely("OnWriteMail", delegate { UiHost.OpenAssistant("compose", null); });
        }

        /* ---------------- 辅助 ---------------- */

        private static string SafeControlId(object control)
        {
            try
            {
                if (control == null) { return ""; }
                dynamic c = control;
                return Convert.ToString(c.Id) ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static object SafeContext(object control)
        {
            try
            {
                if (control == null) { return null; }
                dynamic c = control;
                return c.Context;
            }
            catch
            {
                return null;
            }
        }

        private static void RunSafely(string what, Action action)
        {
            try
            {
                Log.Info("Ribbon 回调: " + what);
                action();
            }
            catch (Exception ex) { Log.Error("Ribbon 回调失败: " + what, ex); }
        }

        private static string LoadResourceText(string resourceName)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException("找不到嵌入资源：" + resourceName);
                }
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
