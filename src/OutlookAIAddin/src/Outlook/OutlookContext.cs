using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OutlookAIAddin
{
    /// <summary>
    /// 当前 Outlook 上下文：阅读窗格 / 邮件列表 / 独立阅读窗口 / 撰写窗口。
    /// 全部通过 COM late binding（dynamic）访问，不依赖 Office PIA / VSTO。
    /// </summary>
    internal sealed class OutlookContext
    {
        public string Kind = "none";     // read | compose | none
        public string Source = "";       // 阅读窗格 / 列表 / 独立窗口（阅读）/ 独立窗口（撰写）
        public string Subject = "";
        public string From = "";
        public string To = "";
        public string BodyText = "";
        public string SelectionText = "";
        public bool WordEditorAvailable;
        public string Error = "";

        internal object Item;
        internal object Inspector;
        internal object WordEditor;

        public bool HasItem
        {
            get { return Kind != "none" && Item != null; }
        }

        public static object AppObject
        {
            get { return Connect.AppObject; }
        }

        /// <summary>
        /// 捕获当前上下文。application 为 null 时使用全局 Outlook.Application；
        /// ribbonContextItem 为右键菜单点击的那封邮件（可为 null）。
        /// </summary>
        public static OutlookContext Capture(object application, object ribbonContextItem)
        {
            return Capture(application, ribbonContextItem, false);
        }

        /// <summary>
        /// preferCompose = true 时，会识别阅读窗格里的“内嵌回复”（inline reply）作为撰写场景。
        /// </summary>
        public static OutlookContext Capture(object application, object ribbonContextItem, bool preferCompose)
        {
            var ctx = new OutlookContext();
            try
            {
                dynamic app = application ?? Connect.AppObject;
                if (app == null)
                {
                    ctx.Error = "Outlook 还没有初始化完成，请稍后再试。";
                    return ctx;
                }

                if (ribbonContextItem != null && IsMailItem(ribbonContextItem))
                {
                    ctx.ReadItem(ribbonContextItem, "阅读窗格 / 列表");
                    return ctx;
                }

                // 用 ActiveWindow 的 Class 判断真正活动的 Outlook 窗口：
                // 34 = Explorer（列表/阅读窗格），35 = Inspector（独立邮件窗口）。
                // 不能用 ActiveInspector()：Explorer 在前台时它仍可能返回后台的 Inspector。
                object activeWindow = TryGetActiveWindow(app);
                int activeClass = GetOutlookClass(activeWindow);
                Log.Info("Capture: activeWindow class=" + activeClass + ", preferCompose=" + preferCompose);

                if (activeClass == 35)
                {
                    object inspector = activeWindow;
                    object item = TryGetCurrentItem(inspector);
                    if (item != null)
                    {
                        bool compose = IsComposeMail(item);
                        ctx.Kind = compose ? "compose" : "read";
                        ctx.Inspector = inspector;
                        ctx.ReadItem(item, compose ? "独立窗口（撰写）" : "独立窗口（阅读）");
                        if (compose) { ctx.ReadWordSelection(); }
                        return ctx;
                    }
                }
                else if (activeClass == 34)
                {
                    object explorer = activeWindow;

                    // 阅读窗格里的“内嵌回复”：挂在 Explorer 上，不在 Inspector 里
                    if (preferCompose)
                    {
                        object inlineItem = TryGetInlineResponse(explorer);
                        if (inlineItem != null && IsMailItem(inlineItem) && !IsSentMail(inlineItem))
                        {
                            ctx.Kind = "compose";
                            ctx.Source = "阅读窗格（内嵌回复）";
                            ctx.ReadItem(inlineItem, ctx.Source);
                            object inlineEditor = TryGetInlineWordEditor(explorer);
                            if (inlineEditor != null) { ctx.ReadWordSelectionFromEditor(inlineEditor); }
                            else { ctx.WordEditorAvailable = false; }
                            return ctx;
                        }
                    }

                    object selected = TryGetFirstSelectedItem(explorer);
                    if (selected != null)
                    {
                        ctx.ReadItem(selected, "阅读窗格 / 列表");
                        return ctx;
                    }
                }

                // 兜底（拿不到 ActiveWindow 时沿用旧逻辑）
                object fallbackInspector = TryGetActiveInspector(app);
                if (fallbackInspector != null)
                {
                    object item = TryGetCurrentItem(fallbackInspector);
                    if (item != null)
                    {
                        bool compose = IsComposeMail(item);
                        ctx.Kind = compose ? "compose" : "read";
                        ctx.Inspector = fallbackInspector;
                        ctx.ReadItem(item, compose ? "独立窗口（撰写）" : "独立窗口（阅读）");
                        if (compose) { ctx.ReadWordSelection(); }
                        return ctx;
                    }
                }

                object fallbackExplorer = TryGetActiveExplorer(app);
                if (fallbackExplorer != null)
                {
                    if (preferCompose)
                    {
                        object inlineItem = TryGetInlineResponse(fallbackExplorer);
                        if (inlineItem != null && IsMailItem(inlineItem) && !IsSentMail(inlineItem))
                        {
                            ctx.Kind = "compose";
                            ctx.Source = "阅读窗格（内嵌回复）";
                            ctx.ReadItem(inlineItem, ctx.Source);
                            object inlineEditor = TryGetInlineWordEditor(fallbackExplorer);
                            if (inlineEditor != null) { ctx.ReadWordSelectionFromEditor(inlineEditor); }
                            else { ctx.WordEditorAvailable = false; }
                            return ctx;
                        }
                    }
                    object item = TryGetFirstSelectedItem(fallbackExplorer);
                    if (item != null)
                    {
                        ctx.ReadItem(item, "阅读窗格 / 列表");
                        return ctx;
                    }
                }

                ctx.Error = "当前没有选中或打开的邮件。请先选中一封邮件（列表或阅读窗格）或打开一封邮件，再试一次。";
            }
            catch (Exception ex)
            {
                ctx.Error = "读取当前邮件失败：" + ex.Message;
                Log.Error("OutlookContext.Capture", ex);
            }
            return ctx;
        }

        private void ReadItem(object item, string source)
        {
            Item = item;
            Source = source;
            if (Kind == "none") { Kind = "read"; }

            dynamic mail = item;
            Subject = SafeString(delegate { return Convert.ToString(mail.Subject); });
            From = SafeString(delegate { return Convert.ToString(mail.SenderName); });
            To = SafeString(delegate { return Convert.ToString(mail.To); });
            BodyText = SafeString(delegate { return Convert.ToString(mail.Body); });
        }

        private void ReadWordSelection()
        {
            try
            {
                dynamic inspector = Inspector;
                object editor = null;
                try { editor = inspector.WordEditor; }
                catch (Exception ex) { Log.Warn("获取 WordEditor 失败：" + ex.Message); }

                if (editor == null)
                {
                    WordEditorAvailable = false;
                    return;
                }
                ReadWordSelectionFromEditor(editor);
            }
            catch (Exception ex)
            {
                WordEditorAvailable = false;
                Log.Warn("读取选中文字失败：" + ex.Message);
            }
        }

        private void ReadWordSelectionFromEditor(object editor)
        {
            try
            {
                if (editor == null)
                {
                    WordEditorAvailable = false;
                    return;
                }
                WordEditor = editor;
                WordEditorAvailable = true;

                dynamic doc = editor;
                dynamic selection = doc.Application.Selection;
                string text = Convert.ToString(selection.Text) ?? "";
                SelectionText = NormalizeWordText(text);
            }
            catch (Exception ex)
            {
                WordEditorAvailable = false;
                Log.Warn("读取选中文字失败（内嵌回复/撰写）: " + ex.Message);
            }
        }

        /// <summary>替换撰写窗口中被选中的文字；返回值 "ok" 表示成功，否则是给人看的错误信息。</summary>
        public string ReplaceSelection(string newText)
        {
            try
            {
                dynamic doc = WordEditor;
                if (doc == null)
                {
                    return "当前撰写窗口无法访问正文编辑器（可能为纯文本格式），无法替换选中文字。可以改用「复制」后手动粘贴。";
                }

                dynamic selection = doc.Application.Selection;
                int selectionType = Convert.ToInt32(selection.Type); // 1 = 光标, 2 = 选中
                if (selectionType == 2)
                {
                    selection.Text = newText ?? "";
                }
                else
                {
                    selection.TypeText(newText ?? "");
                }
                return "ok";
            }
            catch (Exception ex)
            {
                Log.Error("ReplaceSelection", ex);
                return "替换选中文字失败：" + ex.Message;
            }
        }

        /// <summary>在撰写窗口光标处插入文字。</summary>
        public string InsertAtCursor(string text)
        {
            try
            {
                dynamic doc = WordEditor;
                if (doc == null) { return "当前撰写窗口无法访问正文编辑器。"; }
                dynamic selection = doc.Application.Selection;
                selection.TypeText(text ?? "");
                return "ok";
            }
            catch (Exception ex)
            {
                Log.Error("InsertAtCursor", ex);
                return "插入失败：" + ex.Message;
            }
        }

        /// <summary>把当前邮件正文整体替换为指定文字（撰写模式）。</summary>
        public string ReplaceBodyText(string text)
        {
            try
            {
                if (Item == null) { return "没有可写入的邮件。"; }
                dynamic mail = Item;
                mail.Body = text ?? "";
                return "ok";
            }
            catch (Exception ex)
            {
                Log.Error("ReplaceBodyText", ex);
                return "写入正文失败：" + ex.Message;
            }
        }

        /// <summary>把生成好的主题/正文填入当前正在撰写的邮件窗口。</summary>
        public string FillCurrent(string subject, string bodyText, bool replaceBody)
        {
            try
            {
                if (Item == null) { return "当前没有正在编辑的邮件。"; }
                dynamic mail = Item;
                if (!string.IsNullOrWhiteSpace(subject)) { mail.Subject = subject; }
                if (!string.IsNullOrWhiteSpace(bodyText) && replaceBody) { mail.Body = bodyText; }
                return "ok";
            }
            catch (Exception ex)
            {
                Log.Error("FillCurrent", ex);
                return "填入当前草稿失败：" + ex.Message;
            }
        }

        /// <summary>新建一封邮件并显示撰写窗口（而不是发送）。</summary>
        public static string CreateNewMail(string to, string subject, string bodyText)
        {
            try
            {
                dynamic app = Connect.AppObject;
                if (app == null) { return "Outlook 还没有初始化完成。"; }

                dynamic mail = app.CreateItem(0); // olMailItem
                if (!string.IsNullOrWhiteSpace(to)) { mail.To = to; }
                if (!string.IsNullOrWhiteSpace(subject)) { mail.Subject = subject; }
                if (!string.IsNullOrWhiteSpace(bodyText)) { mail.Body = bodyText; }
                mail.Display(false);
                return "ok";
            }
            catch (Exception ex)
            {
                Log.Error("CreateNewMail", ex);
                return "新建邮件失败：" + ex.Message;
            }
        }

        /* ---------------- COM 访问辅助 ---------------- */

        private static object TryGetActiveInspector(object app)
        {
            dynamic a = app;
            try { return a.ActiveInspector(); } catch { }
            try { return a.ActiveInspector; } catch { }
            return null;
        }

        private static object TryGetActiveWindow(object app)
        {
            if (app == null) { return null; }
            dynamic a = app;
            try { return a.ActiveWindow(); } catch { }
            return null;
        }

        private static int GetOutlookClass(object value)
        {
            if (value == null) { return 0; }
            try
            {
                dynamic obj = value;
                return Convert.ToInt32(obj.Class);
            }
            catch
            {
                return 0;
            }
        }

        private static object TryGetInlineResponse(object explorer)
        {
            if (explorer == null) { return null; }
            dynamic e = explorer;
            try { return e.ActiveInlineResponse; } catch { }
            return null;
        }

        private static object TryGetInlineWordEditor(object explorer)
        {
            if (explorer == null) { return null; }
            dynamic e = explorer;
            try { return e.ActiveInlineResponseWordEditor; } catch { }
            return null;
        }

        private static bool IsSentMail(object item)
        {
            if (item == null) { return false; }
            try
            {
                dynamic mail = item;
                return Convert.ToBoolean(mail.Sent);
            }
            catch
            {
                return false;
            }
        }

        private static object TryGetActiveExplorer(object app)
        {
            dynamic a = app;
            try { return a.ActiveExplorer(); } catch { }
            try { return a.ActiveExplorer; } catch { }
            return null;
        }

        private static object TryGetCurrentItem(object inspector)
        {
            dynamic i = inspector;
            try { return i.CurrentItem; } catch { }
            try { return i.CurrentItem(); } catch { }
            return null;
        }

        private static object TryGetFirstSelectedItem(object explorer)
        {
            object selection = null;
            dynamic e = explorer;
            try { selection = e.Selection; } catch { }
            if (selection == null) { return null; }

            dynamic s = selection;
            int count = 0;
            try { count = Convert.ToInt32(s.Count); } catch { return null; }
            if (count <= 0) { return null; }

            try { return s.Item(1); } catch { }
            try { return s[1]; } catch { }
            return null;
        }

        private static bool IsMailItem(object item)
        {
            try
            {
                dynamic mail = item;
                return Convert.ToInt32(mail.Class) == 43; // olMail
            }
            catch
            {
                return false;
            }
        }

        private static bool IsComposeMail(object item)
        {
            try
            {
                dynamic mail = item;
                int itemClass = Convert.ToInt32(mail.Class);
                if (itemClass != 43) { return false; } // olMail
                bool sent = Convert.ToBoolean(mail.Sent);
                return !sent;
            }
            catch (Exception ex)
            {
                Log.Warn("判断撰写/阅读状态失败：" + ex.Message);
                return false;
            }
        }

        private static string SafeString(Func<string> getter)
        {
            try { return getter() ?? ""; }
            catch { return ""; }
        }

        private static string NormalizeWordText(string text)
        {
            if (string.IsNullOrEmpty(text)) { return ""; }
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n')
                .Replace("\u0007", "")   // Word 单元格标记
                .Replace("\u000b", "\n"); // 软换行
            return normalized.TrimEnd('\n');
        }
    }
}
