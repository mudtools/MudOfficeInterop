

//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using System.Drawing;

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示颜色格式的接口，用于定义和管理颜色属性。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointColorFormat : IOfficeObject<IPowerPointColorFormat, MsPowerPoint.ColorFormat>, IDisposable
{
    /// <summary>
    /// 获取创建此颜色格式对象的应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true, NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此颜色格式对象的创建者标识符。
    /// </summary>
    int Creator { get; }

    /// <summary>
    /// 获取此颜色格式对象的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置颜色的RGB值。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    Color RGB { get; set; }

    /// <summary>
    /// 获取颜色类型。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoColorType Type { get; }

    /// <summary>
    /// 获取或设置配色方案中的颜色索引。
    /// </summary>
    PpColorSchemeIndex SchemeColor { get; set; }

    /// <summary>
    /// 获取或设置颜色的色调和阴影值（范围从-1.0到1.0）。
    /// </summary>
    /// <remarks>
    /// 负值表示阴影，正值表示色调。
    /// </remarks>
    float TintAndShade { get; set; }

    /// <summary>
    /// 获取或设置与主题关联的颜色索引。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoThemeColorIndex ObjectThemeColor { get; set; }

    /// <summary>
    /// 获取或设置颜色的亮度值（范围从-1.0到1.0）。
    /// </summary>
    /// <remarks>
    /// 负值表示变暗，正值表示变亮。
    /// </remarks>
    float Brightness { get; set; }
}