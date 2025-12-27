//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 批注集合的封装接口。
/// </summary>
public interface IWordComments : IEnumerable<IWordComment>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取批注数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取批注。
    /// </summary>
    IWordComment this[int index] { get; }

    /// <summary>
    /// 获取第一个批注。
    /// </summary>
    IWordComment First { get; }

    /// <summary>
    /// 获取最后一个批注。
    /// </summary>
    IWordComment Last { get; }

    /// <summary>
    /// 添加新的批注。
    /// </summary>
    /// <param name="range">批注范围。</param>
    /// <param name="text">批注文本。</param>
    /// <param name="author">批注作者。</param>
    /// <returns>新创建的批注。</returns>
    IWordComment Add(IWordRange range, string text = null, string author = null);

    /// <summary>
    /// 删除指定索引的批注。
    /// </summary>
    /// <param name="index">批注索引。</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定范围的批注。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="count">删除数量。</param>
    void DeleteRange(int startIndex, int count);

    /// <summary>
    /// 删除所有批注。
    /// </summary>
    void Clear();

    /// <summary>
    /// 获取所有批注索引列表。
    /// </summary>
    /// <returns>批注索引列表。</returns>
    List<int> GetIndexes();

    /// <summary>
    /// 获取文档总批注字符数。
    /// </summary>
    /// <returns>字符总数。</returns>
    int GetTotalCharacters();

    /// <summary>
    /// 获取文档总批注单词数。
    /// </summary>
    /// <returns>单词总数。</returns>
    int GetTotalWords();

    /// <summary>
    /// 按日期范围筛选批注。
    /// </summary>
    /// <param name="startDate">开始日期。</param>
    /// <param name="endDate">结束日期。</param>
    /// <returns>批注列表。</returns>
    List<IWordComment> GetByDateRange(DateTime startDate, DateTime endDate);

    /// <summary>
    /// 按作者筛选批注。
    /// </summary>
    /// <param name="author">作者名称。</param>
    /// <returns>批注列表。</returns>
    List<IWordComment> GetByAuthor(string author);

    /// <summary>
    /// 查找包含指定文本的批注。
    /// </summary>
    /// <param name="text">要查找的文本。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <param name="matchWholeWord">是否匹配整个单词。</param>
    /// <returns>批注列表。</returns>
    List<IWordComment> FindByText(string text, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 替换所有批注中的指定文本。
    /// </summary>
    /// <param name="findText">要查找的文本。</param>
    /// <param name="replaceText">替换文本。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <param name="matchWholeWord">是否匹配整个单词。</param>
    /// <returns>总替换次数。</returns>
    int ReplaceAllText(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 获取指定范围的批注。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="endIndex">结束索引。</param>
    /// <returns>批注列表。</returns>
    List<IWordComment> GetRange(int startIndex, int endIndex);
}