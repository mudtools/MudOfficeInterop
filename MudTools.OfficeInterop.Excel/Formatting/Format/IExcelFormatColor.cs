
namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// 表示 Excel 中对象的颜色格式设置的封装接口（语义命名为 FormatColor
/// 用于设置前景色、背景色、透明度、主题色等。
/// </summary>
public interface IExcelFormatColor : IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（如 FillFormat、GlowFormat、ShadowFormat 等）。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication Application { get; }

    /// <summary>
    /// 获取或设置颜色的 RGB 值（如 0xFF0000 表示红色）。
    /// 设置此属性会清除主题色设置。
    /// </summary>
    int RGB { get; set; }

    /// <summary>
    /// 获取或设置颜色的主题色类型（如强调色1、背景色等）。
    /// 使用 <see cref="MsoThemeColorIndex"/> 枚举。
    /// 设置此属性会清除 RGB 设置。
    /// </summary>
    MsoThemeColorIndex ThemeColor { get; set; }

    /// <summary>
    /// 获取或设置基于主题色的色调调整（-1.0 ~ 1.0）。
    /// 负值变暗，正值变亮，0 表示不调整。
    /// </summary>
    float TintAndShade { get; set; }

    /// <summary>
    /// 获取或设置颜色的透明度（0-100，0=完全不透明，100=完全透明）。
    /// 内部自动转换为 COM 所需的 0.0~1.0 浮点值。
    /// </summary>
    int Transparency { get; set; }

    /// <summary>
    /// 获取颜色类型（RGB、主题色等）。
    /// 使用 <see cref="MsoColorType"/> 枚举。
    /// </summary>
    MsoColorType Type { get; }
}
