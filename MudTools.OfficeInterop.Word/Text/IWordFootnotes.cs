//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 脚注集合的封装接口。
/// </summary>
public interface IWordFootnotes : IEnumerable<IWordFootnote>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取脚注数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取脚注。
    /// </summary>
    IWordFootnote this[int index] { get; }

    /// <summary>
    /// 获取第一个脚注。
    /// </summary>
    IWordFootnote First { get; }

    /// <summary>
    /// 获取最后一个脚注。
    /// </summary>
    IWordFootnote Last { get; }

    IWordRange Separator { get; }

    IWordRange ContinuationSeparator { get; }

    IWordRange ContinuationNotice { get; }

    /// <summary>
    /// 获取或设置脚注编号方式。
    /// </summary>
    WdNoteNumberStyle NumberStyle { get; set; }

    /// <summary>
    /// 获取或设置脚注起始编号。
    /// </summary>
    int StartingNumber { get; set; }

    /// <summary>
    /// 获取或设置脚注编号格式。
    /// </summary>
    WdNumberingRule NumberingRule { get; set; }

    /// <summary>
    /// 获取或设置脚注位置。
    /// </summary>
    WdFootnoteLocation Location { get; set; }

    /// <summary>
    /// 添加新的脚注。
    /// </summary>
    /// <param name="range">添加脚注的范围。</param>
    /// <param name="referenceText">引用文本。</param>
    /// <param name="noteText">脚注文本。</param>
    /// <returns>新创建的脚注。</returns>
    IWordFootnote Add(IWordRange range, string referenceText = null, string noteText = null);

    /// <summary>
    /// 删除指定索引的脚注。
    /// </summary>
    /// <param name="index">脚注索引。</param>
    void Delete(int index);

    /// <summary>
    /// 删除所有脚注。
    /// </summary>
    void Clear();

    /// <summary>
    /// 获取所有脚注索引列表。
    /// </summary>
    /// <returns>脚注索引列表。</returns>
    List<int> GetIndexes();

    /// <summary>
    /// 重新编号所有脚注。
    /// </summary>
    void Renumber();

    /// <summary>
    /// 获取指定范围内的脚注数量。
    /// </summary>
    /// <param name="range">范围对象。</param>
    /// <returns>脚注数量。</returns>
    int CountInRange(IWordRange range);

    /// <summary>
    /// 查找包含指定文本的脚注。
    /// </summary>
    /// <param name="text">要查找的文本。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <returns>脚注列表。</returns>
    List<IWordFootnote> FindByText(string text, bool matchCase = false);
}