//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 文档节接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordSection : IDisposable
{
    /// <summary>
    /// 获取当前文档归属的<see cref="IWordApplication"/>对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取集合中项的索引号。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示此节是否受到保护以用于表单。
    /// 当设置为 true 时，用户只能在文档的受保护区域中输入数据。
    /// </summary>
    bool ProtectedForForms { get; set; }

    /// <summary>
    /// 获取或设置与节关联的边框集合。
    /// </summary>
    IWordBorders? Borders { get; set; }

    /// <summary>
    /// 获取节范围
    /// </summary>
    IWordRange? Range { get; }

    /// <summary>
    /// 获取表示指定节中所有页眉的集合。
    /// </summary>
    IWordHeadersFooters? Headers { get; }

    /// <summary>
    /// 获取表示指定节中所有页脚的集合。
    /// </summary>
    IWordHeadersFooters? Footers { get; }

    /// <summary>
    /// 获取页面设置
    /// </summary>
    IWordPageSetup? PageSetup { get; set; }

}