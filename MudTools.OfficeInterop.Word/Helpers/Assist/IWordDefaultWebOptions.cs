//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 包含Microsoft Word将文档保存为网页或打开网页时使用的全局应用程序级属性。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordDefaultWebOptions : IOfficeObject<IWordDefaultWebOptions, MsWord.DefaultWebOptions>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个值，指示是否为指定浏览器优化网页。
    /// 当保存为新网页时，针对<see cref="BrowserLevel"/>属性指定的浏览器进行优化；
    /// </summary>
    bool OptimizeForBrowser { get; set; }

    /// <summary>
    /// 获取或设置目标浏览器级别，表示希望针对哪个级别的Web浏览器优化在Microsoft Word中创建的新网页。
    /// </summary>
    WdBrowserLevel BrowserLevel { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在Web浏览器中查看保存的文档时是否使用级联样式表(CSS)进行字体格式化。
    /// 如果为true，Word将创建级联样式表文件并根据<see cref="OrganizeInFolder"/>属性的值将其保存到指定文件夹或与网页相同的文件夹。
    /// 如果为false，则使用HTML &lt;FONT&gt;标签和级联样式表。默认值为true。
    /// </summary>
    bool RelyOnCSS { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示将指定文档保存为网页时，所有支持文件（如背景纹理和图形）是否组织在单独的文件夹中。
    /// 如果为true，支持文件保存在单独文件夹中；如果为false，支持文件保存在与网页相同的文件夹中。默认值为true。
    /// </summary>
    bool OrganizeInFolder { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在将文档保存为网页之前是否自动更新所有支持文件的超链接和路径。
    /// 如果为true，确保在保存文档时链接是最新的；如果为false，则不更新链接。默认值为true。
    /// </summary>
    bool UpdateLinksOnSave { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示将文档保存为网页时是否使用长文件名。
    /// 如果为true，使用长文件名；如果为false，使用DOS文件名格式(8.3)。默认值为true。
    /// </summary>
    bool UseLongFileNames { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示启动Word时是否检查Office应用程序是否为默认HTML编辑器。
    /// 如果为true，Word执行此检查；如果为false，则不执行。默认值为true。
    /// </summary>
    bool CheckIfOfficeIsHTMLEditor { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示启动Word时是否检查Word是否为默认HTML编辑器。
    /// 如果为true，Word执行此检查；如果为false，则不执行。默认值为true。
    /// </summary>
    bool CheckIfWordIsDefaultHTMLEditor { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示将文档保存为网页时是否不从绘图对象生成图像文件。
    /// 如果为true，不生成图像；如果为false，生成图像。默认值为false。
    /// </summary>
    bool RelyOnVML { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示将文档保存为网页时是否允许PNG（便携式网络图形）作为图像格式。
    /// 如果为true，允许PNG作为输出格式；如果为false，不允许。默认值为false。
    /// </summary>
    bool AllowPNG { get; set; }

    /// <summary>
    /// 获取或设置查看保存的文档时建议使用的最小屏幕尺寸（宽度×高度，以像素为单位）。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoScreenSize ScreenSize { get; set; }

    /// <summary>
    /// 获取或设置网页上图形图像和表格单元格的密度（每英寸像素数）。
    /// 设置范围通常为19到480，常用屏幕尺寸的典型设置为72、96和120。
    /// </summary>
    int PixelsPerInch { get; set; }

    /// <summary>
    /// 获取或设置查看保存的文档时Web浏览器要使用的文档编码（代码页或字符集）。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoEncoding Encoding { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示保存网页或纯文本文档时是否使用默认编码，独立于文件打开时的原始编码。
    /// 如果为true，使用默认编码；如果为false，使用文件的原始编码。默认值为false。
    /// </summary>
    bool AlwaysSaveInDefaultEncoding { get; set; }

    /// <summary>
    /// 获取WebPageFonts集合，表示在Word中打开网页且网页中未指定字体信息，
    /// 或当前默认字体无法显示网页中的字符集时，Microsoft Word使用的字体集。
    /// </summary>
    IOfficeWebPageFonts? Fonts { get; }

    /// <summary>
    /// 获取将文档保存为网页、使用长文件名并选择将支持文件保存在单独文件夹时，
    /// Microsoft Word使用的文件夹后缀（即当<see cref="UseLongFileNames"/>和
    /// <see cref="OrganizeInFolder"/>属性都设置为true时）。
    /// </summary>
    string FolderSuffix { get; }

    /// <summary>
    /// 获取或设置在Web浏览器中查看文档时的目标浏览器。
    /// </summary>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoTargetBrowser TargetBrowser { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示Microsoft Word是否使用单文件网页（以前称为Web存档）格式保存新网页。
    /// </summary>
    bool SaveNewWebPagesAsWebArchives { get; set; }
}