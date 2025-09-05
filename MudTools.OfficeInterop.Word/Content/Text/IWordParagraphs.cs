//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 段落集合的封装接口。
/// </summary>
public interface IWordParagraphs : IEnumerable<IWordParagraph>, IDisposable
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
    /// 获取段落数量。
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取段落。
    /// </summary>
    IWordParagraph this[int index] { get; }

    /// <summary>
    /// 获取第一个段落。
    /// </summary>
    IWordParagraph First { get; }

    /// <summary>
    /// 获取最后一个段落。
    /// </summary>
    IWordParagraph Last { get; }

    /// <summary>
    /// 添加新的段落。
    /// </summary>
    /// <param name="text">段落文本。</param>
    /// <param name="beforeParagraph">在指定段落前添加。</param>
    /// <returns>新创建的段落。</returns>
    IWordParagraph Add(string text = null, int beforeParagraph = -1);

    /// <summary>
    /// 在指定位置插入段落。
    /// </summary>
    /// <param name="index">插入位置。</param>
    /// <param name="text">段落文本。</param>
    /// <returns>新插入的段落。</returns>
    IWordParagraph Insert(int index, string text = null);

    /// <summary>
    /// 删除指定索引的段落。
    /// </summary>
    /// <param name="index">段落索引。</param>
    void Delete(int index);

    /// <summary>
    /// 删除指定范围的段落。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="count">删除数量。</param>
    void DeleteRange(int startIndex, int count);

    /// <summary>
    /// 删除所有段落。
    /// </summary>
    void Clear();

    /// <summary>
    /// 获取所有段落索引列表。
    /// </summary>
    /// <returns>段落索引列表。</returns>
    List<int> GetIndexes();

    /// <summary>
    /// 获取文档总字符数。
    /// </summary>
    /// <returns>字符总数。</returns>
    int GetTotalCharacters();

    /// <summary>
    /// 获取文档总单词数。
    /// </summary>
    /// <returns>单词总数。</returns>
    int GetTotalWords();

    /// <summary>
    /// 获取文档总句子数。
    /// </summary>
    /// <returns>句子总数。</returns>
    int GetTotalSentences();

    /// <summary>
    /// 按照指定条件排序段落。
    /// </summary>
    /// <param name="ascending">是否升序。</param>
    void Sort(bool ascending = true);

    /// <summary>
    /// 查找包含指定文本的段落。
    /// </summary>
    /// <param name="text">要查找的文本。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <param name="matchWholeWord">是否匹配整个单词。</param>
    /// <returns>段落列表。</returns>
    List<IWordParagraph> FindByText(string text, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 替换所有段落中的指定文本。
    /// </summary>
    /// <param name="findText">要查找的文本。</param>
    /// <param name="replaceText">替换文本。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <param name="matchWholeWord">是否匹配整个单词。</param>
    /// <returns>总替换次数。</returns>
    int ReplaceAllText(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 获取指定范围的段落。
    /// </summary>
    /// <param name="startIndex">开始索引。</param>
    /// <param name="endIndex">结束索引。</param>
    /// <returns>段落列表。</returns>
    List<IWordParagraph> GetRange(int startIndex, int endIndex);

    /// <summary>
    /// 获取标题段落列表。
    /// </summary>
    /// <returns>标题段落列表。</returns>
    List<IWordParagraph> GetHeadings();

    /// <summary>
    /// 获取空段落列表。
    /// </summary>
    /// <returns>空段落列表。</returns>
    List<IWordParagraph> GetEmptyParagraphs();

    /// <summary>
    /// 批量设置段落格式。
    /// </summary>
    /// <param name="alignment">对齐方式。</param>
    /// <param name="leftIndent">左缩进。</param>
    /// <param name="lineSpacing">行距。</param>
    void SetFormatForAll(WdParagraphAlignment alignment = WdParagraphAlignment.wdAlignParagraphLeft,
                        float leftIndent = 0, float lineSpacing = 1.0f);

    /// <summary>
    /// 获取段落长度统计信息。
    /// </summary>
    /// <returns>长度统计信息。</returns>
    (int MinLength, int MaxLength, double AverageLength) GetLengthStatistics();
}