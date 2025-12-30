//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 EmailSignatureEntry 对象的集合，这些对象表示 Microsoft Word 可用的所有电子邮件签名条目。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordEmailSignatureEntries : IEnumerable<IWordEmailSignatureEntry?>, IOfficeObject<IWordEmailSignatureEntries>, IDisposable
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
    /// 获取指定集合中的项目数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 返回集合中的单个对象。
    /// </summary>
    /// <param name="index">指示序数位置的对象或表示单个对象名称的字符串。</param>
    /// <returns>指定索引处的 EmailSignatureEntry 对象。</returns>
    IWordEmailSignatureEntry? this[int index] { get; }

    /// <summary>
    /// 返回集合中的单个对象。
    /// </summary>
    /// <param name="name">指示序数位置的对象或表示单个对象名称的字符串。</param>
    /// <returns>指定索引处的 EmailSignatureEntry 对象。</returns>
    IWordEmailSignatureEntry? this[string name] { get; }

    /// <summary>
    /// 返回表示新电子邮件签名条目的 EmailSignatureEntry 对象。
    /// </summary>
    /// <param name="name">必需。电子邮件条目的名称。</param>
    /// <param name="range">必需。Range 对象，表示将作为签名添加的文档范围。</param>
    /// <returns>新创建的 EmailSignatureEntry 对象。</returns>
    IWordEmailSignatureEntry? Add(string name, IWordRange range);
}