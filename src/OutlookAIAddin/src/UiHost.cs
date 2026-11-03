using System;

namespace OutlookAIAddin
{
    /// <summary>管理 AI 助手独立窗口（单例；Outlook 退出时关闭）。</summary>
    internal static class UiHost
    {
        private static readonly object Sync = new object();
        private static MainForm _form;

        public static void OpenAssistant(string intent, object item)
        {
            try
            {
                Log.Info("UiHost.OpenAssistant: intent=" + intent + ", item=" + (item != null));
                lock (Sync)
                {
                    if (_form == null || _form.IsDisposed)
                    {
                        _form = new MainForm();
                    }

                    BringToFront(_form);

                    if (item != null)
                    {
                        _form.PrepareFromItem(item, intent);
                    }
                    else
                    {
                        _form.Prepare(intent);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("UiHost.OpenAssistant", ex);
                try
                {
                    System.Windows.Forms.MessageBox.Show(
                        "打开 AI 邮件助手窗口失败：" + ex.Message + "\r\n\r\n详细日志：" + AppPaths.LogFile,
                        "AI 邮件助手", System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                }
                catch { }
            }
        }

        public static void OpenAssistantWithSelection(string intent, string selectionText, object item = null)
        {
            try
            {
                lock (Sync)
                {
                    if (_form == null || _form.IsDisposed)
                    {
                        // 先设置 intent/selection 再 Show：OnFormLoad 里的上下文捕获会用到 intent
                        _form = new MainForm();
                        _form.PrepareWithSelection(intent, selectionText, item);
                        BringToFront(_form);
                    }
                    else
                    {
                        BringToFront(_form);
                        _form.PrepareWithSelection(intent, selectionText, item);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("UiHost.OpenAssistantWithSelection", ex);
            }
        }

        /// <summary>强制把助手窗口带到最前面（即使它已经在后台打开）。</summary>
        private static void BringToFront(MainForm form)
        {
            try
            {
                if (!form.Visible) { form.Show(); }
                if (form.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                {
                    form.WindowState = System.Windows.Forms.FormWindowState.Normal;
                }
                form.Activate();
                form.BringToFront();
                bool topMost = form.TopMost;
                form.TopMost = true;
                form.TopMost = topMost;
                SetForegroundWindow(form.Handle);
            }
            catch (Exception ex)
            {
                Log.Warn("UiHost.BringToFront: " + ex.Message);
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        public static void Reset()
        {
            lock (Sync)
            {
                _form = null;
            }
        }

        public static void CloseAll()
        {
            try
            {
                lock (Sync)
                {
                    if (_form != null && !_form.IsDisposed)
                    {
                        _form.Close();
                        _form.Dispose();
                    }
                    _form = null;
                }
            }
            catch (Exception ex)
            {
                Log.Error("UiHost.CloseAll", ex);
            }
        }
    }
}
