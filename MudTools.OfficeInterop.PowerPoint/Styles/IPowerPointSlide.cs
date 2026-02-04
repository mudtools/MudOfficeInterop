//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// 表示 PowerPoint 演示文稿中的幻灯片。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointSlide : IDisposable
{
    /// <summary>
    /// 获取创建此幻灯片的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此幻灯片的父对象。
    /// </summary>
    /// <value>表示此幻灯片父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取此幻灯片中的形状集合。
    /// </summary>
    /// <value>表示形状集合的 <see cref="IPowerPointShapes"/> 对象。</value>
    IPowerPointShapes? Shapes { get; }

    /// <summary>
    /// 获取此幻灯片的页眉页脚集合。
    /// </summary>
    /// <value>表示页眉页脚集合的 <see cref="IPowerPointHeadersFooters"/> 对象。</value>
    IPowerPointHeadersFooters? HeadersFooters { get; }

    /// <summary>
    /// 获取此幻灯片的幻灯片放映切换效果。
    /// </summary>
    /// <value>表示幻灯片放映切换效果的 <see cref="IPowerPointSlideShowTransition"/> 对象。</value>
    IPowerPointSlideShowTransition? SlideShowTransition { get; }

    /// <summary>
    /// 获取或设置此幻灯片的颜色方案。
    /// </summary>
    /// <value>表示颜色方案的 <see cref="IPowerPointColorScheme"/> 对象。</value>
    IPowerPointColorScheme? ColorScheme { get; set; }

    /// <summary>
    /// 获取此幻灯片的背景形状范围。
    /// </summary>
    /// <value>表示背景形状范围的 <see cref="IPowerPointShapeRange"/> 对象。</value>
    IPowerPointShapeRange? Background { get; }

    /// <summary>
    /// 获取或设置此幻灯片的名称。
    /// </summary>
    /// <value>表示幻灯片名称的字符串。</value>
    string? Name { get; set; }

    /// <summary>
    /// 获取此幻灯片的唯一标识符。
    /// </summary>
    /// <value>表示幻灯片标识符的整数值。</value>
    int SlideID { get; }

    /// <summary>
    /// 获取此幻灯片的打印步骤数。
    /// </summary>
    /// <value>表示打印步骤数的整数值。</value>
    int PrintSteps { get; }

    /// <summary>
    /// 选择此幻灯片。
    /// </summary>
    void Select();

    /// <summary>
    /// 剪切此幻灯片。
    /// </summary>
    void Cut();

    /// <summary>
    /// 复制此幻灯片。
    /// </summary>
    void Copy();

    /// <summary>
    /// 获取或设置此幻灯片的版式。
    /// </summary>
    /// <value>表示幻灯片版式的 <see cref="PpSlideLayout"/> 枚举值。</value>
    PpSlideLayout Layout { get; set; }

    /// <summary>
    /// 复制此幻灯片。
    /// </summary>
    /// <returns>新创建的幻灯片范围。</returns>
    IPowerPointSlideRange? Duplicate();

    /// <summary>
    /// 删除此幻灯片。
    /// </summary>
    void Delete();

    /// <summary>
    /// 获取此幻灯片的标签集合。
    /// </summary>
    /// <value>表示标签集合的 <see cref="IPowerPointTags"/> 对象。</value>
    IPowerPointTags? Tags { get; }

    /// <summary>
    /// 获取此幻灯片在演示文稿中的索引。
    /// </summary>
    /// <value>表示幻灯片索引的整数值。</value>
    int SlideIndex { get; }

    /// <summary>
    /// 获取此幻灯片的编号。
    /// </summary>
    /// <value>表示幻灯片编号的整数值。</value>
    int SlideNumber { get; }

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
    /// 获取此幻灯片的备注页。
    /// </summary>
    /// <value>表示备注页的 <see cref="IPowerPointSlideRange"/> 对象。</value>
    IPowerPointSlideRange? NotesPage { get; }

    /// <summary>
    /// 获取此幻灯片使用的母版。
    /// </summary>
    /// <value>表示母版的 <see cref="IPowerPointMaster"/> 对象。</value>
    IPowerPointMaster? Master { get; }

    /// <summary>
    /// 获取此幻灯片中的超链接集合。
    /// </summary>
    /// <value>表示超链接集合的 <see cref="IPowerPointHyperlinks"/> 对象。</value>
    IPowerPointHyperlinks? Hyperlinks { get; }

    /// <summary>
    /// 导出此幻灯片为指定格式的图像文件。
    /// </summary>
    /// <param name="fileName">导出文件的名称。</param>
    /// <param name="filterName">导出过滤器的名称。</param>
    /// <param name="scaleWidth">导出图像的宽度缩放比例。</param>
    /// <param name="scaleHeight">导出图像的高度缩放比例。</param>
    void Export(string fileName, string filterName, int scaleWidth = 0, int scaleHeight = 0);

    /// <summary>
    /// 获取此幻灯片中的脚本集合。
    /// </summary>
    /// <value>表示脚本集合的 <see cref="IOfficeScripts"/> 对象。</value>
    IOfficeScripts? Scripts { get; }

    /// <summary>
    /// 获取此幻灯片中的注释集合。
    /// </summary>
    /// <value>表示注释集合的 <see cref="IPowerPointComments"/> 对象。</value>
    IPowerPointComments? Comments { get; }

    /// <summary>
    /// 获取或设置此幻灯片的设计模板。
    /// </summary>
    /// <value>表示设计模板的 <see cref="IPowerPointDesign"/> 对象。</value>
    IPowerPointDesign? Design { get; set; }

    /// <summary>
    /// 将此幻灯片移动到指定位置。
    /// </summary>
    /// <param name="toPos">要移动到的目标位置索引。</param>
    void MoveTo(int toPos);

    /// <summary>
    /// 获取此幻灯片的时间线对象。
    /// </summary>
    /// <value>表示时间线对象的 <see cref="IPowerPointTimeLine"/> 对象。</value>
    IPowerPointTimeLine? TimeLine { get; }

    /// <summary>
    /// 应用指定的模板到此幻灯片。
    /// </summary>
    /// <param name="fileName">模板文件的名称。</param>
    void ApplyTemplate(string fileName);

    /// <summary>
    /// 获取此幻灯片所在的节编号。
    /// </summary>
    /// <value>表示节编号的整数值。</value>
    int SectionNumber { get; }

    /// <summary>
    /// 获取或设置此幻灯片的自定义版式。
    /// </summary>
    /// <value>表示自定义版式的 <see cref="IPowerPointCustomLayout"/> 对象。</value>
    IPowerPointCustomLayout? CustomLayout { get; set; }

    /// <summary>
    /// 应用指定的主题到此幻灯片。
    /// </summary>
    /// <param name="themeName">主题名称。</param>
    void ApplyTheme(string themeName);

    /// <summary>
    /// 获取此幻灯片的主题颜色方案。
    /// </summary>
    /// <value>表示主题颜色方案的 <see cref="IOfficeThemeColorScheme"/> 对象。</value>
    IOfficeThemeColorScheme? ThemeColorScheme { get; }

    /// <summary>
    /// 应用指定的主题颜色方案到此幻灯片。
    /// </summary>
    /// <param name="themeColorSchemeName">主题颜色方案的名称。</param>
    void ApplyThemeColorScheme(string themeColorSchemeName);

    /// <summary>
    /// 获取或设置此幻灯片的背景样式。
    /// </summary>
    /// <value>表示背景样式的 <see cref="MsoBackgroundStyleIndex"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoBackgroundStyleIndex BackgroundStyle { get; set; }

    /// <summary>
    /// 获取此幻灯片的客户数据。
    /// </summary>
    /// <value>表示客户数据的 <see cref="IPowerPointCustomerData"/> 对象。</value>
    IPowerPointCustomerData? CustomerData { get; }

    /// <summary>
    /// 将此幻灯片发布到幻灯片库。
    /// </summary>
    /// <param name="slideLibraryUrl">幻灯片库的 URL。</param>
    /// <param name="overwrite">指示是否覆盖现有幻灯片的布尔值。</param>
    /// <param name="useSlideOrder">指示是否使用幻灯片顺序的布尔值。</param>
    void PublishSlides(string slideLibraryUrl, bool overwrite = false, bool useSlideOrder = false);

    /// <summary>
    /// 将此幻灯片移动到指定节的开始位置。
    /// </summary>
    /// <param name="toSection">目标节的索引。</param>
    void MoveToSectionStart(int toSection);

    /// <summary>
    /// 获取此幻灯片在节中的索引。
    /// </summary>
    /// <value>表示节索引的整数值。</value>
    [ComPropertyWrap(PropertyName = "sectionIndex")]
    int SectionIndex { get; }

    /// <summary>
    /// 获取一个值，指示此幻灯片是否有备注页。
    /// </summary>
    /// <value>指示是否有备注页的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HasNotesPage { get; }

    /// <summary>
    /// 应用指定的模板到此幻灯片（增强版本）。
    /// </summary>
    /// <param name="fileName">模板文件的名称。</param>
    /// <param name="variantGUID">变体 GUID。</param>
    void ApplyTemplate2(string fileName, string variantGUID);
}