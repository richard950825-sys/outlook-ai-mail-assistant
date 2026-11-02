using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace OutlookAIAddin
{
    internal static class AppPaths
    {
        public static string RootDir
        {
            get
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "OutlookAIAssistant");
                Directory.CreateDirectory(dir);
                return dir;
            }
        }

        public static string SettingsFile { get { return Path.Combine(RootDir, "settings.json"); } }
        public static string TemplatesFile { get { return Path.Combine(RootDir, "templates.json"); } }
        public static string LogFile { get { return Path.Combine(RootDir, "addin.log"); } }
    }

    internal static class Log
    {
        private static readonly object Sync = new object();

        public static void Info(string message) { Write("INFO", message); }
        public static void Warn(string message) { Write("WARN", message); }

        public static void Error(string context, Exception ex)
        {
            if (ex == null)
            {
                Write("ERROR", context);
                return;
            }
            Write("ERROR", context + " | " + ex.GetType().Name + ": " + ex.Message
                + Environment.NewLine + "  " + (ex.StackTrace ?? "").Replace(Environment.NewLine, Environment.NewLine + "  "));
        }

        public static void Error(string message) { Write("ERROR", message); }

        private static void Write(string level, string message)
        {
            try
            {
                string path = AppPaths.LogFile;
                lock (Sync)
                {
                    var info = new FileInfo(path);
                    if (info.Exists && info.Length > 2 * 1024 * 1024)
                    {
                        try { File.Delete(path + ".bak"); } catch { }
                        try { File.Move(path, path + ".bak"); } catch { }
                    }
                    File.AppendAllText(
                        path,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " [" + level + "] " + message + Environment.NewLine,
                        Encoding.UTF8);
                }
            }
            catch
            {
                // 日志失败不能影响加载项
            }
        }
    }

    internal static class Json
    {
        private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer
        {
            MaxJsonLength = int.MaxValue,
            RecursionLimit = 100
        };

        public static string Serialize(object value)
        {
            return Serializer.Serialize(value);
        }

        public static T Deserialize<T>(string value)
        {
            return Serializer.Deserialize<T>(value);
        }
    }

    internal static class Dpapi
    {
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("OutlookAIAddin.Settings.v1");

        public static string Protect(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) { return ""; }
            try
            {
                byte[] data = ProtectedData.Protect(
                    Encoding.UTF8.GetBytes(plainText), Entropy, DataProtectionScope.CurrentUser);
                return Convert.ToBase64String(data);
            }
            catch (Exception ex)
            {
                Log.Error("Dpapi.Protect", ex);
                return "";
            }
        }

        public static string Unprotect(string encrypted)
        {
            if (string.IsNullOrEmpty(encrypted)) { return ""; }
            try
            {
                byte[] data = ProtectedData.Unprotect(
                    Convert.FromBase64String(encrypted), Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(data);
            }
            catch (Exception ex)
            {
                Log.Error("Dpapi.Unprotect", ex);
                return "";
            }
        }
    }
}
