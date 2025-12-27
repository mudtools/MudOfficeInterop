//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Word 句子集合的封装接口。
/// </summary>
public interface IWordSentences : IEnumerable<IWordRange>, IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取句子数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取句子范围。
    /// </summary>
    IWordRange this[int index] { get; }

    /// <summary>
    /// 获取第一个句子。
    /// </summary>
    IWordRange First { get; }

    /// <summary>
    /// 获取最后一个句子。
    /// </summary>
    IWordRange Last { get; }

    /// <summary>
    /// 获取指定范围的句子。
    /// </summary>
    /// <param name="start">开始位置。</param>
    /// <param name="end">结束位置。</param>
    /// <returns>句子范围对象。</returns>
    IWordRange GetRange(int start, int end);


    /// <summary>
    /// 查找指定文本。
    /// </summary>
    /// <param name="findText">要查找的文本。</param>
    /// <param name="forward">是否向前查找。</param>
    /// <param name="wrap">是否循环查找。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <param name="matchWholeWord">是否匹配整个单词。</param>
    /// <param name="matchWildcards">是否使用通配符。</param>
    /// <param name="matchSoundsLike">是否匹配发音。</param>
    /// <param name="matchAllWordForms">是否匹配所有词形。</param>
    /// <returns>找到的句子范围，未找到返回null。</returns>
    IWordRange Find(string findText, bool forward = true, WdFindWrap wrap = WdFindWrap.wdFindStop,
                   bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false,
                   bool matchSoundsLike = false, bool matchAllWordForms = false);

    /// <summary>
    /// 替换指定文本。
    /// </summary>
    /// <param name="findText">要查找的文本。</param>
    /// <param name="replaceText">替换文本。</param>
    /// <param name="replaceAll">是否替换所有匹配项。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <param name="matchWholeWord">是否匹配整个单词。</param>
    /// <param name="matchWildcards">是否使用通配符。</param>
    /// <param name="matchSoundsLike">是否匹配发音。</param>
    /// <param name="matchAllWordForms">是否匹配所有词形。</param>
    /// <returns>替换的次数。</returns>
    int Replace(string findText, string replaceText, bool replaceAll = false,
                bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false,
                bool matchSoundsLike = false, bool matchAllWordForms = false);

    /// <summary>
    /// 添加文本到末尾。
    /// </summary>
    /// <param name="text">要添加的文本。</param>
    /// <returns>新添加的句子范围。</returns>
    IWordRange Add(string text);

    /// <summary>
    /// 插入文本到指定位置。
    /// </summary>
    /// <param name="index">插入位置。</param>
    /// <param name="text">要插入的文本。</param>
    /// <returns>新插入的句子范围。</returns>
    IWordRange Insert(int index, string text);

    /// <summary>
    /// 删除指定范围的句子。
    /// </summary>
    /// <param name="start">开始位置。</param>
    /// <param name="count">删除句子数。</param>
    void Delete(int start, int count);

    /// <summary>
    /// 清除所有句子内容。
    /// </summary>
    void Clear();

    /// <summary>
    /// 获取所有句子索引列表。
    /// </summary>
    /// <returns>句子索引列表。</returns>
    List<int> GetIndexes();

    /// <summary>
    /// 获取文本内容。
    /// </summary>
    /// <returns>完整文本内容。</returns>
    string GetText();

    /// <summary>
    /// 设置文本内容。
    /// </summary>
    /// <param name="text">新的文本内容。</param>
    void SetText(string text);

    /// <summary>
    /// 获取指定范围的子句子集合。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="length">长度。</param>
    /// <returns>子句子集合范围。</returns>
    IWordRange Substring(int startIndex, int length);

    /// <summary>
    /// 检查是否包含指定文本。
    /// </summary>
    /// <param name="text">要检查的文本。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <returns>是否包含。</returns>
    bool Contains(string text, bool matchCase = false);

    /// <summary>
    /// 统计句子中的单词数。
    /// </summary>
    /// <param name="index">句子索引。</param>
    /// <returns>单词数量。</returns>
    int GetWordCount(int index);
}