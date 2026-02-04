//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 表示 PowerPoint 中的超链接对象。
/// </summary>
[ComObjectWrap(ComNamespace = "MsPowerPoint")]
public interface IPowerPointHyperlink : IOfficeObject<IPowerPointHyperlink, MsPowerPoint.Hyperlink>, IDisposable
{
    /// <summary>
    /// 获取创建此超链接的 PowerPoint 应用程序实例。
    /// </summary>
    /// <value>表示 PowerPoint 应用程序的 <see cref="IPowerPointApplication"/> 对象。</value>
    [ComPropertyWrap(NeedDispose = false)]
    IPowerPointApplication? Application { get; }

    /// <summary>
    /// 获取此超链接的父对象。
    /// </summary>
    /// <value>表示此超链接父对象的 <see cref="object"/>。</value>
    object? Parent { get; }

    /// <summary>
    /// 获取此超链接的类型。
    /// </summary>
    /// <value>表示超链接类型的 <see cref="MsoHyperlinkType"/> 枚举值。</value>
    [ComPropertyWrap(ComNamespace = "MsCore")]
    MsoHyperlinkType Type { get; }

    /// <summary>
    /// 获取或设置超链接的地址。
    /// </summary>
    /// <value>表示超链接地址的字符串。</value>
    string? Address { get; set; }

    /// <summary>
    /// 获取或设置超链接的子地址。
    /// </summary>
    /// <value>表示超链接子地址的字符串。</value>
    string? SubAddress { get; set; }

    /// <summary>
    /// 将此超链接添加到收藏夹。
    /// </summary>
    void AddToFavorites();

    /// <summary>
    /// 获取或设置电子邮件超链接的主题。
    /// </summary>
    /// <value>表示电子邮件主题的字符串。</value>
    string? EmailSubject { get; set; }

    /// <summary>
    /// 获取或设置超链接的屏幕提示文本。
    /// </summary>
    /// <value>表示屏幕提示文本的字符串。</value>
    string? ScreenTip { get; set; }

    /// <summary>
    /// 获取或设置超链接的显示文本。
    /// </summary>
    /// <value>表示显示文本的字符串。</value>
    string? TextToDisplay { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在打开超链接后是否返回到原始文档。
    /// </summary>
    /// <value>指示是否返回的布尔值。</value>
    [ComPropertyWrap(NeedConvert = true)]
    bool ShowAndReturn { get; set; }

    /// <summary>
    /// 跟踪此超链接。
    /// </summary>
    void Follow();

    /// <summary>
    /// 创建新文档作为此超链接的目标。
    /// </summary>
    /// <param name="fileName">要创建的新文档的文件名。</param>
    /// <param name="editNow">指示是否立即编辑新文档的布尔值。</param>
    /// <param name="overwrite">指示是否覆盖现有文件的布尔值。</param>
    void CreateNewDocument(string fileName, [ConvertTriState] bool editNow, [ConvertTriState] bool overwrite);

    /// <summary>
    /// 删除此超链接。
    /// </summary>
    void Delete();
}
