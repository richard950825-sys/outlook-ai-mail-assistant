using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace OutlookAIAddin
{
    /// <summary>把嵌入的 PNG 转成 Ribbon 需要的 IPictureDisp。</summary>
    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter() : base("") { }

        public static object GetImage(string resourceName)
        {
            try
            {
                using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (stream == null) { return null; }
                    using (Image image = Image.FromStream(stream))
                    {
                        return GetIPictureDispFromPicture(image);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warn("加载 Ribbon 图标失败：" + resourceName + " | " + ex.Message);
                return null;
            }
        }
    }
}
