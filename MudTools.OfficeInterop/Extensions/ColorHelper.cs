//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop;

internal static class ColorHelper
{
    /// <summary>
    /// 将Com颜色转换为System.Drawing.Color
    /// </summary>
    public static Color ConvertToColor(int color)
    {
        // 处理无色和自动颜色
        if (color == -4142 || color == 0xFFFFFFFF)
        {
            return Color.Transparent;
        }

        return ColorTranslator.FromOle(color);
    }

    /// <summary>
    /// 将Com颜色转换为System.Drawing.Color
    /// </summary>
    public static Color ConvertToColor(object colorObj)
    {
        if (colorObj == null)
            return Color.Transparent;

        try
        {
            // 检查是否为整型（Excel中常见的颜色表示）
            if (colorObj is int intColor)
            {
                // 处理Excel的特殊颜色值
                if (intColor == -4142)  // xlColorIndexNone
                    return Color.Transparent;
                if (intColor == -4105)  // xlColorIndexAutomatic
                    return Color.Transparent; // 或返回默认颜色

                return ColorTranslator.FromOle(intColor);
            }

            // 如果是double类型（有时Excel返回double）
            if (colorObj is double doubleColor)
            {
                return ColorTranslator.FromOle((int)doubleColor);
            }

            // 如果是Color对象直接返回
            if (colorObj is Color color)
            {
                return color;
            }

            return Color.Transparent;
        }
        catch
        {
            return Color.Transparent;
        }
    }

    /// <summary>
    /// 将System.Drawing.Color转换为Excel颜色
    /// </summary>
    public static int ConvertToComColor(Color color)
    {
        if (color == Color.Transparent)
        {
            return -4142;  // Excel无色常量
        }

        return ColorTranslator.ToOle(color);
    }
}
