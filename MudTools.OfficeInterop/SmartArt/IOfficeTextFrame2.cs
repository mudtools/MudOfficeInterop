//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 表示 Office 中文本框架的接口封装，提供对文本框的各种属性和操作的访问。
/// </summary>
[ComObjectWrap(ComNamespace = "MsCore")]
public interface IOfficeTextFrame2 : IOfficeObject<IOfficeTextFrame2>, IDisposable
{
    /// <summary>
    /// 获取文本框架的父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置文本框架底部边距。
    /// </summary>
    float MarginBottom { get; set; }

    /// <summary>
    /// 获取或设置文本框架左侧边距。
    /// </summary>
    float MarginLeft { get; set; }

    /// <summary>
    /// 获取或设置文本框架右侧边距。
    /// </summary>
    float MarginRight { get; set; }

    /// <summary>
    /// 获取或设置文本框架顶部边距。
    /// </summary>
    float MarginTop { get; set; }

    /// <summary>
    /// 获取文本框架中的文本范围
    /// </summary>
    IOfficeTextRange2? TextRange { get; }

    /// <summary>
    /// 获取文本框架中的文本列
    /// </summary>
    IOfficeTextColumn2? Column { get; }

    /// <summary>
    /// 获取文本框架中的标尺
    /// </summary>
    IOfficeRuler2? Ruler { get; }

    /// <summary>
    /// 获取或设置文本方向。
    /// </summary>
    MsoTextOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置文本的水平对齐方式。
    /// </summary>
    MsoHorizontalAnchor HorizontalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本的垂直对齐方式。
    /// </summary>
    MsoVerticalAnchor VerticalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本路径格式。
    /// </summary>
    MsoPathFormat PathFormat { get; set; }

    /// <summary>
    /// 获取或设置文本变形格式。
    /// </summary>
    MsoWarpFormat WarpFormat { get; set; }

    /// <summary>
    /// 获取或设置预设文本效果。
    /// </summary>
    MsoPresetTextEffect WordArtformat { get; set; }

    /// <summary>
    /// 获取或设置是否自动换行。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool WordWrap { get; set; }

    /// <summary>
    /// 获取文本框架是否包含文本。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasText { get; }

    /// <summary>
    /// 获取或设置文本是否不旋转。
    /// </summary>
    [ComPropertyWrap(NeedConvert = true)]
    bool NoTextRotation { get; set; }

    /// <summary>
    /// 获取或设置文本框架的自动调整大小行为。
    /// </summary>
    MsoAutoSize AutoSize { get; set; }

    /// <summary>
    /// 获取文本框架的三维格式设置。
    /// </summary>
    IOfficeThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 删除文本框架中的所有文本。
    /// </summary>
    void DeleteText();
}