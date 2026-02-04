//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示形状的文本框架，提供对文本格式、布局和效果的访问。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointTextFrame2 : IOfficeObject<IPowerPointTextFrame2, MsPowerPoint.TextFrame2>, IDisposable
{
    /// <summary>
    /// 获取创建此文本框架的应用程序。
    /// </summary>
    /// <value>表示应用程序的对象。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此对象的应用程序的创建者代码。
    /// </summary>
    /// <value>创建者标识符。</value>
    int Creator { get; }

    /// <summary>
    /// 获取文本框架的父对象。
    /// </summary>
    /// <value>父对象，通常是形状。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置文本框架的下边距（磅）。
    /// </summary>
    /// <value>文本框架的下边距（磅）。</value>
    float MarginBottom { get; set; }

    /// <summary>
    /// 获取或设置文本框架的左边距（磅）。
    /// </summary>
    /// <value>文本框架的左边距（磅）。</value>
    float MarginLeft { get; set; }

    /// <summary>
    /// 获取或设置文本框架的右边距（磅）。
    /// </summary>
    /// <value>文本框架的右边距（磅）。</value>
    float MarginRight { get; set; }

    /// <summary>
    /// 获取或设置文本框架的上边距（磅）。
    /// </summary>
    /// <value>文本框架的上边距（磅）。</value>
    float MarginTop { get; set; }

    /// <summary>
    /// 获取或设置文本框架中文本的方向。
    /// </summary>
    /// <value>文本的方向。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoTextOrientation Orientation { get; set; }

    /// <summary>
    /// 获取或设置文本框架的水平对齐锚点。
    /// </summary>
    /// <value>文本的水平对齐方式。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoHorizontalAnchor HorizontalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本框架的垂直对齐锚点。
    /// </summary>
    /// <value>文本的垂直对齐方式。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoVerticalAnchor VerticalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本框架的路径格式。
    /// </summary>
    /// <value>文本沿路径排列的格式。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPathFormat PathFormat { get; set; }

    /// <summary>
    /// 获取或设置文本框架的弯曲格式。
    /// </summary>
    /// <value>文本的弯曲效果格式。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoWarpFormat WarpFormat { get; set; }

    /// <summary>
    /// 获取或设置文本框架的艺术字格式。
    /// </summary>
    /// <value>文本的艺术字效果格式。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoPresetTextEffect WordArtFormat { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示文本是否自动换行。
    /// </summary>
    /// <value>如果文本自动换行，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool WordWrap { get; set; }

    /// <summary>
    /// 获取或设置文本框架的自动调整大小方式。
    /// </summary>
    /// <value>文本框架的自动调整大小设置。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoAutoSize AutoSize { get; set; }

    /// <summary>
    /// 获取文本框架的三维格式设置。
    /// </summary>
    /// <value>文本框架的三维格式对象。</value>
    IPowerPointThreeDFormat? ThreeD { get; }

    /// <summary>
    /// 获取一个值，指示文本框架是否包含文本。
    /// </summary>
    /// <value>如果文本框架包含文本，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasText { get; }

    /// <summary>
    /// 获取文本框架中的文本范围。
    /// </summary>
    /// <value>文本框架中的文本范围对象。</value>
    IOfficeTextRange2? TextRange { get; }

    /// <summary>
    /// 获取文本框架的列设置。
    /// </summary>
    /// <value>文本框架的列格式对象。</value>
    IOfficeTextColumn2? Column { get; }

    /// <summary>
    /// 获取文本框架的标尺设置。
    /// </summary>
    /// <value>文本框架的标尺对象。</value>
    IOfficeRuler2? Ruler { get; }

    /// <summary>
    /// 删除文本框架中的所有文本。
    /// </summary>
    void DeleteText();

    /// <summary>
    /// 获取或设置一个值，指示文本在形状旋转时是否保持不旋转。
    /// </summary>
    /// <value>如果文本不随形状旋转，则为 <see langword="true"/>；否则为 <see langword="false"/>。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool NoTextRotation { get; set; }
}