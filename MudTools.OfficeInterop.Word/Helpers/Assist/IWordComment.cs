//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Word 批注的封装接口。
/// </summary>
public interface IWordComment : IDisposable
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
    /// 获取批注索引。
    /// </summary>
    int Index { get; }

    /// <summary>
    /// 获取或设置批注作者。
    /// </summary>
    string Author { get; set; }

    /// <summary>
    /// 获取或设置批注初始。
    /// </summary>
    string Initial { get; set; }

    /// <summary>
    /// 获取批注范围。
    /// </summary>
    IWordRange Range { get; }

    /// <summary>
    /// 获取批注文本范围。
    /// </summary>
    IWordRange CommentRange { get; }

    /// <summary>
    /// 获取批注日期时间。
    /// </summary>
    DateTime Date { get; }

    /// <summary>
    /// 获取批注字符数。
    /// </summary>
    int CharactersCount { get; }

    /// <summary>
    /// 获取批注单词数。
    /// </summary>
    int WordsCount { get; }

    /// <summary>
    /// 获取批注句子数。
    /// </summary>
    int SentencesCount { get; }

    /// <summary>
    /// 选择批注。
    /// </summary>
    void Select();

    /// <summary>
    /// 删除批注。
    /// </summary>
    void Delete();

    /// <summary>
    /// 复制批注。
    /// </summary>
    void Copy();

    /// <summary>
    /// 获取批注文本内容。
    /// </summary>
    /// <returns>批注文本。</returns>
    string GetText();

    /// <summary>
    /// 设置批注文本内容。
    /// </summary>
    /// <param name="text">文本内容。</param>
    void SetText(string text);

    /// <summary>
    /// 追加文本到批注末尾。
    /// </summary>
    /// <param name="text">要追加的文本。</param>
    void AppendText(string text);

    /// <summary>
    /// 检查批注是否包含指定文本。
    /// </summary>
    /// <param name="text">要检查的文本。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <returns>是否包含。</returns>
    bool ContainsText(string text, bool matchCase = false);

    /// <summary>
    /// 查找并替换批注中的文本。
    /// </summary>
    /// <param name="findText">要查找的文本。</param>
    /// <param name="replaceText">替换文本。</param>
    /// <param name="matchCase">是否匹配大小写。</param>
    /// <param name="matchWholeWord">是否匹配整个单词。</param>
    /// <returns>替换次数。</returns>
    int ReplaceText(string findText, string replaceText, bool matchCase = false, bool matchWholeWord = false);

    /// <summary>
    /// 获取批注引用文本。
    /// </summary>
    /// <returns>引用文本。</returns>
    string GetReferenceText();

    /// <summary>
    /// 获取批注引用范围。
    /// </summary>
    /// <returns>引用范围。</returns>
    IWordRange GetReferenceRange();
}