//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示对 Microsoft Word AutoCorrectEntries 集合的封装接口。
/// 包含所有自动更正条目，支持按名称或索引访问、添加、删除等操作。
/// </summary>
[ComCollectionWrap(ComNamespace = "MsWord")]
public interface IWordAutoCorrectEntries : IEnumerable<IWordAutoCorrectEntry?>, IOfficeObject<IWordAutoCorrectEntries, MsWord.AutoCorrectEntries>, IDisposable
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
    /// 获取自动更正条目的总数。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过 1-based 索引获取指定位置的自动更正条目。
    /// </summary>
    /// <param name="index">索引（从1开始）。</param>
    /// <returns>封装后的条目对象，若不存在则返回 null。</returns>
    IWordAutoCorrectEntry? this[int index] { get; }

    /// <summary>
    /// 通过名称（触发词）获取指定的自动更正条目。
    /// </summary>
    /// <param name="name">条目的名称（如 "teh"）。</param>
    /// <returns>封装后的条目对象，若不存在则返回 null。</returns>
    IWordAutoCorrectEntry? this[string name] { get; }

    /// <summary>
    /// 添加一个新的自动更正条目。
    /// </summary>
    /// <param name="name">触发词（如 "omw"）。</param>
    /// <param name="value">替换内容（如 "on my way"）。</param>
    /// <returns>新创建的封装条目对象。</returns>
    /// <exception cref="ArgumentException">当 name 或 value 为空时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加失败时抛出。</exception>
    IWordAutoCorrectEntry? Add(string name, string value);

    /// <summary>
    /// 添加一个新的自动更正条目，并使用富文本作为替换内容。
    /// </summary>
    /// <param name="name">触发词（如 "omw"）。</param>
    /// <param name="range">包含替换内容的富文本范围。</param>
    /// <returns>新创建的封装条目对象。</returns>
    /// <exception cref="ArgumentException">当 name 或 range 为空时抛出。</exception>
    /// <exception cref="InvalidOperationException">当添加失败时抛出。</exception>
    IWordAutoCorrectEntry? AddRichText(string name, IWordRange range);
}