//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.Excel;
/// <summary>
/// 表示 Excel 中对象的颜色格式设置的封装接口（语义命名为 FormatColor
/// 用于设置前景色、背景色、透明度、主题色等。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelFormatColor : IOfficeObject<IExcelFormatColor, MsExcel.FormatColor>, IDisposable
{
    /// <summary>
    /// 获取此对象的父对象（如 FillFormat、GlowFormat、ShadowFormat 等）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IExcelApplication? Application { get; }

    /// <summary>
    /// 获取或设置颜色。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color Color { get; set; }

    /// <summary>
    /// 获取或设置颜色索引类型。
    /// 使用 <see cref="XlColorIndex"/> 枚举，可以设置为自动颜色或无颜色。
    /// </summary>
    XlColorIndex ColorIndex { get; set; }

    /// <summary>
    /// 获取或设置颜色的主题色类型（如强调色1、背景色等）。
    /// 使用 <see cref="MsoThemeColorIndex"/> 枚举。
    /// 设置此属性会清除 RGB 设置。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore", NeedConvert = true)]
    MsoThemeColorIndex ThemeColor { get; set; }

    /// <summary>
    /// 获取或设置基于主题色的色调调整（-1.0 ~ 1.0）。
    /// 负值变暗，正值变亮，0 表示不调整。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    float TintAndShade { get; set; }
}
