using System;
using System.IO;

namespace OutlookAIAddin
{
    /// <summary>
    /// 调查开关：%APPDATA%\OutlookAIAssistant\menu-mode.txt 内容为 native 时，
    /// 不拦截右键（放行 Word 原生菜单），仅做诊断记录。
    /// </summary>
    internal static class DiagnosticsMode
    {
        public static bool IsNativeMenuProbe()
        {
            try
            {
                string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "OutlookAIAssistant", "menu-mode.txt");
                if (!File.Exists(path)) { return false; }
                string text = File.ReadAllText(path).Trim().ToLowerInvariant();
                return text.Contains("native");
            }
            catch
            {
                return false;
            }
        }
    }
}
