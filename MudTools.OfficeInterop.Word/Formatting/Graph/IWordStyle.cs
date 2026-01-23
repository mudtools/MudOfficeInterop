//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Style 的接口，用于操作文档样式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordStyle : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置表格的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "BaseStyle", NeedConvert = true)]
    IWordStyle? BaseStyle { get; set; }

    /// <summary>
    /// 获取或设置表格的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "BaseStyle", NeedConvert = true)]
    string? BaseStyleName { get; set; }

    /// <summary>
    /// 获取或设置表格的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "BaseStyle", NeedConvert = true)]
    WdBuiltinStyle BaseStyleType { get; set; }

    /// <summary>
    /// 获取一个值，该值指示此样式当前是否正在使用中。
    /// </summary>
    bool InUse { get; }

    /// <summary>
    /// 获取样式的本地化名称。
    /// </summary>
    string NameLocal { get; }

    /// <summary>
    /// 获取或设置样式的类型。
    /// </summary>
    WdStyleType Type { get; }

    /// <summary>
    /// 获取或设置是否自动更新样式。
    /// </summary>
    bool AutomaticallyUpdate { get; set; }

    /// <summary>
    /// 获取或设置是否为快捷样式。
    /// </summary>
    bool QuickStyle { get; set; }

    /// <summary>
    /// 获取或设置是否可见。
    /// </summary>
    bool Visibility { get; set; }

    /// <summary>
    /// 获取样式的字体格式封装对象。
    /// </summary>
    IWordFont? Font { get; }

    /// <summary>
    /// 获取样式的段落格式封装对象。
    /// </summary>
    IWordParagraphFormat? ParagraphFormat { get; }

    /// <summary>
    /// 获取样式的编号格式封装对象。
    /// </summary>
    IWordListTemplate? ListTemplate { get; }

    /// <summary>
    /// 删除此样式。
    /// </summary>
    void Delete();
}