using System;
using System.IO;
using System.Text;

namespace OutlookAIAddin
{
    /// <summary>保存到 %APPDATA%\OutlookAIAssistant\settings.json 的设置。</summary>
    public sealed class SettingsData
    {
        public string baseUrl = "https://api.deepseek.com/v1";
        public string model = "deepseek-chat";
        public string apiKeyEnc = "";
        public string summaryLanguage = "zh";    // zh = 中文摘要；source = 跟随原文语言
        public string translateTarget = "auto";  // auto = 中↔英自动；en = 译成英文；zh = 译成中文
        public double temperature = 0.3;
    }

    internal static class SettingsStore
    {
        private static readonly object Sync = new object();
        private static SettingsData _current;

        public static SettingsData Current
        {
            get
            {
                lock (Sync)
                {
                    if (_current == null) { _current = LoadCore(); }
                    return _current;
                }
            }
        }

        public static SettingsData Reload()
        {
            lock (Sync)
            {
                _current = LoadCore();
                return _current;
            }
        }

        private static SettingsData LoadCore()
        {
            var data = new SettingsData();
            try
            {
                if (File.Exists(AppPaths.SettingsFile))
                {
                    string text = File.ReadAllText(AppPaths.SettingsFile, Encoding.UTF8);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        var loaded = Json.Deserialize<SettingsData>(text);
                        if (loaded != null) { data = loaded; }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("读取设置失败，使用默认设置", ex);
            }

            Normalize(data);
            return data;
        }

        private static void Normalize(SettingsData data)
        {
            if (data == null) { return; }
            if (string.IsNullOrWhiteSpace(data.baseUrl)) { data.baseUrl = "https://api.deepseek.com/v1"; }
            data.baseUrl = data.baseUrl.Trim().TrimEnd('/');
            if (string.IsNullOrWhiteSpace(data.model)) { data.model = "deepseek-chat"; }
            if (string.IsNullOrWhiteSpace(data.summaryLanguage)) { data.summaryLanguage = "zh"; }
            if (string.IsNullOrWhiteSpace(data.translateTarget)) { data.translateTarget = "auto"; }
            if (data.temperature < 0 || data.temperature > 1) { data.temperature = 0.3; }
        }

        public static void Save(SettingsData data)
        {
            lock (Sync)
            {
                Normalize(data);
                Directory.CreateDirectory(AppPaths.RootDir);
                File.WriteAllText(AppPaths.SettingsFile, Json.Serialize(data), new UTF8Encoding(false));
                _current = data;
            }
        }

        public static string GetApiKey()
        {
            return Dpapi.Unprotect(Current.apiKeyEnc);
        }

        public static void SetApiKey(string apiKey)
        {
            Current.apiKeyEnc = Dpapi.Protect(apiKey ?? "");
        }
    }
}
