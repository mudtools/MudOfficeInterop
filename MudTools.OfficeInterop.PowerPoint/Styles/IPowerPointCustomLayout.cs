//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// 表示 PowerPoint 中的自定义版式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointCustomLayout : IOfficeObject<IPowerPointCustomLayout, MsPowerPoint.CustomLayout>, IDisposable
{
    /// <summary>
    /// 获取创建此自定义版式的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此自定义版式的父对象。
    /// </summary>
    /// <value>表示此自定义版式父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取此自定义版式中的形状集合。
    /// </summary>
    /// <value>表示形状集合的 <see cref="IPowerPointShapes"/> 对象。</value>
    IPowerPointShapes? Shapes { get; }

    /// <summary>
    /// 获取此自定义版式的页眉页脚集合。
    /// </summary>
    /// <value>表示页眉页脚集合的 <see cref="IPowerPointHeadersFooters"/> 对象。</value>
    IPowerPointHeadersFooters? HeadersFooters { get; }

    /// <summary>
    /// 获取此自定义版式的背景形状范围。
    /// </summary>
    /// <value>表示背景形状范围的 <see cref="IPowerPointShapeRange"/> 对象。</value>
    IPowerPointShapeRange? Background { get; }

    /// <summary>
    /// 获取或设置此自定义版式的名称。
    /// </summary>
    /// <value>表示自定义版式名称的字符串。</value>
    string? Name { get; set; }

    /// <summary>
    /// 删除此自定义版式。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取此自定义版式的高度（以磅为单位）。
    /// </summary>
    /// <value>表示高度的浮点数。</value>
    float Height { get; }

    /// <summary>
    /// 获取此自定义版式的宽度（以磅为单位）。
    /// </summary>
    /// <value>表示宽度的浮点数。</value>
    float Width { get; }

    /// <summary>
    /// 获取此自定义版式中的超链接集合。
    /// </summary>
    /// <value>表示超链接集合的 <see cref="IPowerPointHyperlinks"/> 对象。</value>
    IPowerPointHyperlinks? Hyperlinks { get; }

    /// <summary>
    /// 获取与此自定义版式关联的设计对象。
    /// </summary>
    /// <value>表示设计对象的 <see cref="IPowerPointDesign"/> 对象。</value>
    IPowerPointDesign? Design { get; }

    /// <summary>
    /// 获取此自定义版式的时间线对象。
    /// </summary>
    /// <value>表示时间线对象的 <see cref="IPowerPointTimeLine"/> 对象。</value>
    IPowerPointTimeLine? TimeLine { get; }

    /// <summary>
    /// 获取此自定义版式的幻灯片放映切换效果。
    /// </summary>
    /// <value>表示幻灯片放映切换效果的 <see cref="IPowerPointSlideShowTransition"/> 对象。</value>
    IPowerPointSlideShowTransition? SlideShowTransition { get; }

    /// <summary>
    /// 获取或设置此自定义版式的匹配名称。
    /// </summary>
    /// <value>表示匹配名称的字符串。</value>
    string? MatchingName { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示此自定义版式是否被保留。
    /// </summary>
    /// <value>指示版式是否被保留的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Preserved { get; set; }

    /// <summary>
    /// 获取此自定义版式在集合中的索引。
    /// </summary>
    /// <value>表示索引的整数值。</value>
    int Index { get; }

    /// <summary>
    /// 选择此自定义版式。
    /// </summary>
    void Select();

    /// <summary>
    /// 剪切此自定义版式。
    /// </summary>
    void Cut();

    /// <summary>
    /// 复制此自定义版式。
    /// </summary>
    void Copy();

    /// <summary>
    /// 复制此自定义版式。
    /// </summary>
    /// <returns>新创建的自定义版式副本。</returns>
    IPowerPointCustomLayout? Duplicate();

    /// <summary>
    /// 将此自定义版式移动到指定位置。
    /// </summary>
    /// <param name="toPos">要移动到的目标位置索引。</param>
    void MoveTo(int toPos);

    /// <summary>
    /// 获取或设置一个值，指示是否显示母版形状。
    /// </summary>
    /// <value>指示是否显示母版形状的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool DisplayMasterShapes { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否跟随母版背景。
    /// </summary>
    /// <value>指示是否跟随母版背景的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool FollowMasterBackground { get; set; }

    /// <summary>
    /// 获取此自定义版式的主题颜色方案。
    /// </summary>
    /// <value>表示主题颜色方案的 <see cref="IOfficeThemeColorScheme"/> 对象。</value>
    IOfficeThemeColorScheme? ThemeColorScheme { get; }

    /// <summary>
    /// 获取此自定义版式的客户数据。
    /// </summary>
    /// <value>表示客户数据的 <see cref="IPowerPointCustomerData"/> 对象。</value>
    IPowerPointCustomerData? CustomerData { get; }

    /// <summary>
    /// 获取此自定义版式的参考线集合。
    /// </summary>
    /// <value>表示参考线集合的 <see cref="IPowerPointGuides"/> 对象。</value>
    IPowerPointGuides? Guides { get; }
}