using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;
using System.Windows.Automation.Text;

namespace OutlookAIAddin
{
    /// <summary>
    /// 为 Ribbon 原生正文菜单读取选区。保留原有类名供撰写菜单回调使用；
    /// 不再替换 WndProc、不拦截右键，也不创建替代菜单或贴边窗口。
    /// </summary>
    internal static class WordContextMenuHook
    {
        /// <summary>在打开助手之前，捕获本次阅读菜单对应的邮件和选区。</summary>
        internal static string CaptureReadSelection(object ribbonContext, out object item)
        {
            item = null;
            try
            {
                dynamic app = Connect.AppObject;
                if (app == null) { return ""; }

                // ActiveWindow 返回实际活动的 Explorer/Inspector，避免使用后台 Inspector。
                object source = ribbonContext;
                int sourceClass = GetOutlookClass(source);
                if (sourceClass != 34 && sourceClass != 35)
                {
                    source = app.ActiveWindow();
                    sourceClass = GetOutlookClass(source);
                }
                dynamic window = source;
                if (sourceClass == 35) // olInspector
                {
                    item = window.CurrentItem;
                }
                else if (sourceClass == 34) // olExplorer
                {
                    // 内嵌回复属于撰写场景，不能借阅读菜单处理另一封邮件。
                    try { if (window.ActiveInlineResponse != null) { return ""; } } catch { }
                    dynamic selection = window.Selection;
                    if (Convert.ToInt32(selection.Count) != 1) { return ""; }
                    item = selection.Item(1);
                }
                if (GetOutlookClass(item) != 43) { item = null; return ""; }
                dynamic mail = item;
                if (!Convert.ToBoolean(mail.Sent)) { item = null; return ""; }

                // 独立阅读窗口有 WordEditor；读取该文档的选区，不激活其他窗口。
                if (sourceClass == 35)
                {
                    try
                    {
                        dynamic document = window.WordEditor;
                        string text = NormalizeWordText(Convert.ToString(document.Windows[1].Selection.Text));
                        if (text.Length > 0) { return text; }
                    }
                    catch (Exception ex) { Log.Warn("读取原生阅读窗口选区: " + ex.Message); }
                }

                // 阅读窗格不公开 WordEditor，沿用已验证的 UIA TextPattern。
                // 只检查本次活动 Outlook 窗口内的可见正文，绝不搜索其他邮件窗口。
                IntPtr root = IntPtr.Zero;
                try
                {
                    var oleWindow = source as IOleWindow;
                    if (oleWindow != null) { oleWindow.GetWindow(out root); }
                }
                catch { }
                if (root == IntPtr.Zero) { root = GetAncestor(GetForegroundWindow(), 3); } // GA_ROOTOWNER
                uint pid;
                GetWindowThreadProcessId(root, out pid);
                if (pid != (uint)System.Diagnostics.Process.GetCurrentProcess().Id) { return ""; }

                string selected = "";
                EnumChildWindows(root, delegate (IntPtr child, IntPtr unused)
                {
                    var name = new StringBuilder(64);
                    GetClassName(child, name, name.Capacity);
                    if (IsWindowVisible(child) && name.ToString() == "_WwG")
                    {
                        selected = GetUiaSelectionText(child);
                        if (selected.Length > 0) { return false; }
                    }
                    return true;
                }, IntPtr.Zero);
                return selected;
            }
            catch (Exception ex)
            {
                Log.Warn("CaptureReadSelection: " + ex.Message);
                item = null;
                return "";
            }
        }

        /// <summary>保持撰写窗口原有选区读取行为。</summary>
        internal static string CaptureComposeSelection()
        {
            try
            {
                dynamic app = Connect.AppObject;
                if (app == null) { return ""; }

                // 1) 独立撰写窗口
                dynamic inspector = null;
                try { inspector = app.ActiveInspector(); } catch { }
                if (inspector != null)
                {
                    try { inspector.Activate(); } catch { }
                    dynamic doc = null;
                    try { doc = inspector.WordEditor; } catch { }
                    if (doc != null)
                    {
                        try { doc.Activate(); } catch { }
                        dynamic selection = doc.Application.Selection;
                        return NormalizeWordText(Convert.ToString(selection.Text) ?? "");
                    }
                }

                // 2) 阅读窗格里的内嵌回复（inline reply）
                dynamic explorer = null;
                try { explorer = app.ActiveExplorer(); } catch { }
                if (explorer != null)
                {
                    try
                    {
                        object inlineItem = explorer.ActiveInlineResponse;
                        object editor = explorer.ActiveInlineResponseWordEditor;
                        if (inlineItem != null && editor != null)
                        {
                            dynamic inlineDoc = editor;
                            dynamic selection = inlineDoc.Application.Selection;
                            return NormalizeWordText(Convert.ToString(selection.Text) ?? "");
                        }
                    }
                    catch (Exception ex) { Log.Warn("读取内嵌回复选区: " + ex.Message); }
                }
                return "";
            }
            catch (Exception ex)
            {
                Log.Warn("CaptureComposeSelection: " + ex.Message);
                return "";
            }
        }

        private static int GetOutlookClass(object value)
        {
            if (value == null) { return 0; }
            try { dynamic obj = value; return Convert.ToInt32(obj.Class); }
            catch { return 0; }
        }

        private static string GetUiaSelectionText(IntPtr hwnd)
        {
            try
            {
                AutomationElement element = AutomationElement.FromHandle(hwnd);
                object patternObject;
                if (element == null || !element.TryGetCurrentPattern(TextPattern.Pattern, out patternObject)) { return ""; }
                var pattern = patternObject as TextPattern;
                if (pattern == null) { return ""; }
                TextPatternRange[] ranges = pattern.GetSelection();
                if (ranges == null) { return ""; }
                var text = new StringBuilder();
                foreach (TextPatternRange range in ranges) { text.Append(range.GetText(-1)); }
                return NormalizeWordText(text.ToString());
            }
            catch (Exception ex)
            {
                Log.Warn("GetUiaSelectionText: " + ex.Message);
                return "";
            }
        }

        private static string NormalizeWordText(string text)
        {
            return (text ?? "").Replace("\r\n", "\n").Replace('\r', '\n')
                .Replace("\u0007", "").Trim();
        }

        [ComImport, Guid("00000114-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleWindow
        {
            [PreserveSig] int GetWindow(out IntPtr hwnd);
            [PreserveSig] int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool enterMode);
        }

        private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hwnd, int flags);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr hwnd, EnumWindowsProc callback, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hwnd, StringBuilder className, int maxCount);
    }
}
