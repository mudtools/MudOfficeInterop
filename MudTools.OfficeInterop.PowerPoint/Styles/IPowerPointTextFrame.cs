//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;
/// <summary>
/// 表示 PowerPoint 形状中的文本框。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointTextFrame : IOfficeObject<IPowerPointTextFrame, MsPowerPoint.TextFrame>, IDisposable
{
    /// <summary>
    /// 获取创建此文本框的应用程序实例。
    /// </summary>
    /// <value>表示应用程序的 <see cref="IPowerPointApplication"/>。</value>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取创建此文本框的应用程序的创建者代码。
    /// </summary>
    /// <value>表示创建者代码的整数值。</value>
    int Creator { get; }

    /// <summary>
    /// 获取此文本框的父对象。
    /// </summary>
    /// <value>表示此文本框父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置文本框的底部边距（以磅为单位）。
    /// </summary>
    /// <value>表示底部边距的浮点数。</value>
    float MarginBottom { get; set; }

    /// <summary>
    /// 获取或设置文本框的左侧边距（以磅为单位）。
    /// </summary>
    /// <value>表示左侧边距的浮点数。</value>
    float MarginLeft { get; set; }

    /// <summary>
    /// 获取或设置文本框的右侧边距（以磅为单位）。
    /// </summary>
    /// <value>表示右侧边距的浮点数。</value>
    float MarginRight { get; set; }

    /// <summary>
    /// 获取或设置文本框的顶部边距（以磅为单位）。
    /// </summary>
    /// <value>表示顶部边距的浮点数。</value>
    float MarginTop { get; set; }

    /// <summary>
    /// 获取或设置文本的方向。
    /// </summary>
    /// <value>表示文本方向的 <see cref="MsoTextOrientation"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoTextOrientation Orientation { get; set; }

    /// <summary>
    /// 获取一个值，指示文本框是否包含文本。
    /// </summary>
    /// <value>指示是否包含文本的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasText { get; }

    /// <summary>
    /// 获取文本框的文本范围对象。
    /// </summary>
    /// <value>表示文本范围的 <see cref="IPowerPointTextRange"/> 对象。</value>
    IPowerPointTextRange? TextRange { get; }

    /// <summary>
    /// 获取文本框的标尺设置。
    /// </summary>
    /// <value>表示标尺的 <see cref="IPowerPointRuler"/> 对象。</value>
    IPowerPointRuler? Ruler { get; }

    /// <summary>
    /// 获取或设置文本的水平对齐方式。
    /// </summary>
    /// <value>表示水平对齐方式的 <see cref="MsoHorizontalAnchor"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoHorizontalAnchor HorizontalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本的垂直对齐方式。
    /// </summary>
    /// <value>表示垂直对齐方式的 <see cref="MsoVerticalAnchor"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoVerticalAnchor VerticalAnchor { get; set; }

    /// <summary>
    /// 获取或设置文本框的自动调整大小方式。
    /// </summary>
    /// <value>表示自动调整大小方式的 <see cref="PpAutoSize"/> 枚举值。</value>
    PpAutoSize AutoSize { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示文本是否自动换行。
    /// </summary>
    /// <value>指示是否自动换行的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool WordWrap { get; set; }

    /// <summary>
    /// 删除文本框中的所有文本。
    /// </summary>
    void DeleteText();
}