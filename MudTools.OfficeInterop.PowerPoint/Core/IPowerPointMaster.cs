//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 的幻灯片母版。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointMaster : IDisposable
{
    /// <summary>
    /// 获取创建此幻灯片母版的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此幻灯片母版的父对象。
    /// </summary>
    /// <value>表示此幻灯片母版父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取此幻灯片母版中的形状集合。
    /// </summary>
    /// <value>表示形状集合的 <see cref="IPowerPointShapes"/> 对象。</value>
    IPowerPointShapes? Shapes { get; }

    /// <summary>
    /// 获取此幻灯片母版的页眉页脚集合。
    /// </summary>
    /// <value>表示页眉页脚集合的 <see cref="IPowerPointHeadersFooters"/> 对象。</value>
    IPowerPointHeadersFooters? HeadersFooters { get; }

    /// <summary>
    /// 获取或设置此幻灯片母版的颜色方案。
    /// </summary>
    /// <value>表示颜色方案的 <see cref="IPowerPointColorScheme"/> 对象。</value>
    IPowerPointColorScheme? ColorScheme { get; set; }

    /// <summary>
    /// 获取此幻灯片母版的背景形状范围。
    /// </summary>
    /// <value>表示背景形状范围的 <see cref="IPowerPointShapeRange"/> 对象。</value>
    IPowerPointShapeRange? Background { get; }

    /// <summary>
    /// 获取或设置此幻灯片母版的名称。
    /// </summary>
    /// <value>表示幻灯片母版名称的字符串。</value>
    string? Name { get; set; }

    /// <summary>
    /// 删除此幻灯片母版。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取此幻灯片母版的高度（以磅为单位）。
    /// </summary>
    /// <value>表示高度的浮点数。</value>
    float Height { get; }

    /// <summary>
    /// 获取此幻灯片母版的宽度（以磅为单位）。
    /// </summary>
    /// <value>表示宽度的浮点数。</value>
    float Width { get; }

    /// <summary>
    /// 获取此幻灯片母版的文本样式集合。
    /// </summary>
    /// <value>表示文本样式集合的 <see cref="IPowerPointTextStyles"/> 对象。</value>
    IPowerPointTextStyles? TextStyles { get; }

    /// <summary>
    /// 获取此幻灯片母版中的超链接集合。
    /// </summary>
    /// <value>表示超链接集合的 <see cref="IPowerPointHyperlinks"/> 对象。</value>
    IPowerPointHyperlinks? Hyperlinks { get; }

    /// <summary>
    /// 获取此幻灯片母版中的脚本集合。
    /// </summary>
    /// <value>表示脚本集合的 <see cref="IOfficeScripts"/> 对象。</value>
    IOfficeScripts? Scripts { get; }

    /// <summary>
    /// 获取此幻灯片母版的设计模板。
    /// </summary>
    /// <value>表示设计模板的 <see cref="IPowerPointDesign"/> 对象。</value>
    IPowerPointDesign? Design { get; }

    /// <summary>
    /// 获取此幻灯片母版的时间线对象。
    /// </summary>
    /// <value>表示时间线对象的 <see cref="IPowerPointTimeLine"/> 对象。</value>
    IPowerPointTimeLine? TimeLine { get; }

    /// <summary>
    /// 获取此幻灯片母版的幻灯片放映切换效果。
    /// </summary>
    /// <value>表示幻灯片放映切换效果的 <see cref="IPowerPointSlideShowTransition"/> 对象。</value>
    IPowerPointSlideShowTransition? SlideShowTransition { get; }

    /// <summary>
    /// 获取此幻灯片母版的自定义版式集合。
    /// </summary>
    /// <value>表示自定义版式集合的 <see cref="IPowerPointCustomLayouts"/> 对象。</value>
    IPowerPointCustomLayouts? CustomLayouts { get; }

    /// <summary>
    /// 获取此幻灯片母版的 Office 主题。
    /// </summary>
    /// <value>表示 Office 主题的 <see cref="IOfficeOfficeTheme"/> 对象。</value>
    IOfficeOfficeTheme? Theme { get; }

    /// <summary>
    /// 应用指定的主题到此幻灯片母版。
    /// </summary>
    /// <param name="themeName">主题名称。</param>
    void ApplyTheme(string themeName);

    /// <summary>
    /// 获取或设置此幻灯片母版的背景样式。
    /// </summary>
    /// <value>表示背景样式的 <see cref="MsoBackgroundStyleIndex"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBackgroundStyleIndex BackgroundStyle { get; set; }

    /// <summary>
    /// 获取此幻灯片母版的客户数据。
    /// </summary>
    /// <value>表示客户数据的 <see cref="IPowerPointCustomerData"/> 对象。</value>
    IPowerPointCustomerData? CustomerData { get; }

    /// <summary>
    /// 获取此幻灯片母版的参考线集合。
    /// </summary>
    /// <value>表示参考线集合的 <see cref="IPowerPointGuides"/> 对象。</value>
    IPowerPointGuides? Guides { get; }
}