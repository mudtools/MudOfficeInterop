//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Range 的接口，用于操作 Word 文档中的文本范围。
/// </summary>
public interface IWordRange : IDisposable
{
    /// <summary>
    /// 获取当前文档归属的<see cref="IWordApplication"/>对象。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取或设置范围内的文本内容。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取范围的起始位置。
    /// </summary>
    int Start { get; set; }

    /// <summary>
    /// 获取范围的结束位置。
    /// </summary>
    int End { get; set; }

    /// <summary>
    /// 获取范围的字体样式封装对象。
    /// </summary>
    IWordFont? Font { get; }

    /// <summary>
    /// 获取范围的段落格式封装对象。
    /// </summary>
    IWordParagraphFormat? ParagraphFormat { get; }

    IWordBorders? Borders { get; }

    IWordCharacters? Characters { get; }

    IWordFootnotes? Footnotes { get; }

    IWordEndnotes? Endnotes { get; }

    IWordTables? Tables { get; }

    IWordSections? Sections { get; }

    IWordSentences? Sentences { get; }

    IWordWords? Words { get; }

    IWordParagraphs? Paragraphs { get; }

    IWordShading? Shading { get; }

    IWordFields? Fields { get; }

    IWordBookmarks? Bookmarks { get; }

    IWordListFormat? ListFormat { get; }

    /// <summary>
    /// 获取范围内的字符数。
    /// </summary>
    int CharactersCount { get; }

    /// <summary>
    /// 获取范围内的单词数。
    /// </summary>
    int WordsCount { get; }

    /// <summary>
    /// 获取范围内的段落数。
    /// </summary>
    int ParagraphsCount { get; }

    /// <summary>
    /// 在范围后插入文本。
    /// </summary>
    /// <param name="text">要插入的文本。</param>
    /// <returns>新插入文本的范围。</returns>
    void InsertAfter(string text);

    /// <summary>
    /// 在范围前插入文本。
    /// </summary>
    /// <param name="text">要插入的文本。</param>
    /// <returns>新插入文本的范围。</returns>
    void InsertBefore(string text);

    /// <summary>
    /// 在范围后插入换行符。
    /// </summary>
    void InsertParagraphAfter();

    /// <summary>
    /// 在范围前插入换行符。
    /// </summary>
    void InsertParagraphBefore();

    /// <summary>
    /// 删除范围内的文本。
    /// </summary>
    void Delete();

    /// <summary>
    /// 选择此范围（使光标定位到此范围）。
    /// </summary>
    void Select();

    /// <summary>
    /// 复制范围内容到剪贴板。
    /// </summary>
    void Copy();

    /// <summary>
    /// 从剪贴板粘贴内容到此范围。
    /// </summary>
    void Paste();

    /// <summary>
    /// 查找并替换范围内的文本。
    /// </summary>
    /// <param name="findText">要查找的文本。</param>
    /// <param name="replaceText">替换的文本。</param>
    /// <param name="matchCase">是否区分大小写。</param>
    /// <param name="matchWholeWord">是否匹配整个单词。</param>
    /// <returns>找到并替换的次数。</returns>
    bool FindAndReplace(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 获取指定索引的字符范围。
    /// </summary>
    /// <param name="index">字符索引（从1开始）。</param>
    /// <returns>字符范围。</returns>
    IWordRange GetCharacter(int index);

    /// <summary>
    /// 获取指定索引的单词范围。
    /// </summary>
    /// <param name="index">单词索引（从1开始）。</param>
    /// <returns>单词范围。</returns>
    IWordRange GetWord(int index);

    /// <summary>
    /// 获取指定索引的段落范围。
    /// </summary>
    /// <param name="index">段落索引（从1开始）。</param>
    /// <returns>段落范围。</returns>
    IWordRange GetParagraph(int index);
}