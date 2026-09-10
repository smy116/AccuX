using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace AccuX.AddIn.Ribbon
{
    /// <summary>
    /// Ribbon 图标加载（规格 §7）。
    /// <para>
    /// 图标资源统一由 AccuX.AddIn 管理，通过嵌入资源加载；
    /// 加载失败只返回 null，绝不影响 Command 本身执行。
    /// </para>
    /// </summary>
    internal static class RibbonImageProvider
    {
        private const string ResourcePrefix = "AccuX.AddIn.Resources.";

        /// <summary>
        /// 按控件 ID 返回 IPictureDisp 图标；失败返回 null。
        /// </summary>
        public static object GetImage(string controlId)
        {
            try
            {
                var name = ResourcePrefix + MapFileName(controlId);
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
                {
                    if (stream == null)
                    {
                        return null;
                    }

                    using (var image = Image.FromStream(stream))
                    {
                        // 转换为 IPictureDisp 供 Office 使用；AxHost 无需 stdole 编译期引用。
                        return AxHostConverter.Convert(image);
                    }
                }
            }
            catch
            {
                // 图标失败不得影响 Command 执行。
                return null;
            }
        }

        private static string MapFileName(string controlId)
        {
            switch (controlId)
            {
                case "accux.basic.round":
                    return "round.png";
                case "accux.basic.convert":
                    return "convert.png";
                case "accux.basic.sum":
                    return "sum.png";
                case "accux.basic.uppercase":
                    return "uppercase.png";
                case "accux.basic.directory":
                    return "directory.png";
                case "accux.basic.comment":
                    return "comment.png";
                case "accux.mark.green":
                    return "mark-green.png";
                case "accux.mark.red":
                    return "mark-red.png";
                case "accux.mark.yellow":
                    return "mark-yellow.png";
                case "accux.mark.blue":
                    return "mark-blue.png";
                case "accux.compare.exists":
                    return "sum.png";
                default:
                    return "round.png";
            }
        }
    }

    /// <summary>
    /// 使用 Windows Forms AxHost 将 Image 转换为 IPictureDisp。
    /// 派生类只是为暴露受保护的静态转换方法。
    /// </summary>
    internal sealed class AxHostConverter : AxHost
    {
        private AxHostConverter() : base(string.Empty)
        {
        }

        public static object Convert(Image image)
        {
            return GetIPictureDispFromPicture(image);
        }
    }
}
