//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 文档中列表段落的集合。
/// </summary>
public interface IWordListParagraphs : IEnumerable<IWordParagraph>, IDisposable
{
    /// <summary>
    /// 获取与该对象关联的应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取集合中的段落数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取列表段落。
    /// </summary>
    /// <param name="index">从 1 开始的索引。</param>
    IWordParagraph this[int index] { get; }
}
