using System;
using System.Globalization;
using AccuX.Core.Operations;

namespace AccuX.Host
{
    /// <summary>
    /// 宿主侧纯值转换辅助方法：不持有 Application 状态、不访问宿主对象，可直接单元测试。
    /// 供各能力实现类与 AccuX.Core.Tests 共用。
    /// </summary>
    internal static class ExcelComHelper
    {
        /// <summary>
        /// 将 #RRGGBB 文本转换为 Excel/WPS 的 OLE 颜色值（BGR）。
        /// </summary>
        internal static int ParseOleColor(string hexColor)
        {
            if (string.IsNullOrWhiteSpace(hexColor)
                || hexColor.Length != 7
                || hexColor[0] != '#')
            {
                throw new HostOperationException("颜色值无效，必须使用 #RRGGBB 格式。");
            }

            try
            {
                var red = int.Parse(hexColor.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                var green = int.Parse(hexColor.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                var blue = int.Parse(hexColor.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                return red + (green << 8) + (blue << 16);
            }
            catch (FormatException ex)
            {
                throw new HostOperationException("颜色值无效，必须使用 #RRGGBB 格式。", ex);
            }
            catch (OverflowException ex)
            {
                throw new HostOperationException("颜色值无效，必须使用 #RRGGBB 格式。", ex);
            }
        }
    }
}
