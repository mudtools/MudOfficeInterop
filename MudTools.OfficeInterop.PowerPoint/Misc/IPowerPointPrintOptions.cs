//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！


namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 演示文稿的打印选项。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointPrintOptions : IOfficeObject<IPowerPointPrintOptions, MsPowerPoint.PrintOptions>, IDisposable
{
    /// <summary>
    /// 获取创建此打印选项的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取或设置打印颜色类型。
    /// </summary>
    /// <value>表示打印颜色类型的 <see cref="PpPrintColorType"/> 枚举值。</value>
    PpPrintColorType PrintColorType { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否逐份打印。
    /// </summary>
    /// <value>指示是否逐份打印的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool Collate { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否缩放以适应页面。
    /// </summary>
    /// <value>指示是否缩放以适应页面的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool FitToPage { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否为幻灯片添加边框。
    /// </summary>
    /// <value>指示是否为幻灯片添加边框的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool FrameSlides { get; set; }

    /// <summary>
    /// 获取或设置打印份数。
    /// </summary>
    /// <value>表示打印份数的整数值。</value>
    int NumberOfCopies { get; set; }

    /// <summary>
    /// 获取或设置打印输出类型。
    /// </summary>
    /// <value>表示打印输出类型的 <see cref="PpPrintOutputType"/> 枚举值。</value>
    PpPrintOutputType OutputType { get; set; }

    /// <summary>
    /// 获取此打印选项的父对象。
    /// </summary>
    /// <value>表示此打印选项父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否打印隐藏的幻灯片。
    /// </summary>
    /// <value>指示是否打印隐藏幻灯片的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool PrintHiddenSlides { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否在后台打印。
    /// </summary>
    /// <value>指示是否在后台打印的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool PrintInBackground { get; set; }

    /// <summary>
    /// 获取或设置打印范围类型。
    /// </summary>
    /// <value>表示打印范围类型的 <see cref="PpPrintRangeType"/> 枚举值。</value>
    PpPrintRangeType RangeType { get; set; }

    /// <summary>
    /// 获取打印范围集合。
    /// </summary>
    /// <value>表示打印范围集合的 <see cref="IPowerPointPrintRanges"/> 对象。</value>
    IPowerPointPrintRanges? Ranges { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否将字体作为图形打印。
    /// </summary>
    /// <value>指示是否将字体作为图形打印的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool PrintFontsAsGraphics { get; set; }

    /// <summary>
    /// 获取或设置幻灯片放映的名称。
    /// </summary>
    /// <value>表示幻灯片放映名称的字符串。</value>
    string? SlideShowName { get; set; }

    /// <summary>
    /// 获取或设置活动打印机的名称。
    /// </summary>
    /// <value>表示活动打印机名称的字符串。</value>
    string? ActivePrinter { get; set; }

    /// <summary>
    /// 获取或设置讲义的打印顺序。
    /// </summary>
    /// <value>表示讲义打印顺序的 <see cref="PpPrintHandoutOrder"/> 枚举值。</value>
    PpPrintHandoutOrder HandoutOrder { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否打印注释。
    /// </summary>
    /// <value>指示是否打印注释的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool PrintComments { get; set; }

    /// <summary>
    /// 获取或设置节的索引。
    /// </summary>
    /// <value>表示节索引的整数值。</value>
    [ComPropertyWrap(PropertyName = "sectionIndex")]
    int SectionIndex { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否使用高质量打印。
    /// </summary>
    /// <value>指示是否使用高质量打印的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool HighQuality { get; set; }
}