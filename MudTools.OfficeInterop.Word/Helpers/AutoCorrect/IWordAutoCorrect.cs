//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;
/// <summary>
/// 表示 Microsoft Word 中的自动更正功能。
/// <para>注：使用 Application.AutoCorrect 或 Application.AutoCorrectEmail 属性可返回 AutoCorrect 对象。</para>
/// </summary>
public interface IWordAutoCorrect : IDisposable
{
    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    #region 自动更正选项属性 (AutoCorrect Options Properties)

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Word 是否自动更正你在键入时无意中使用 CAPS LOCK 键的实例。
    /// </summary>
    bool CorrectCapsLock { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Word 是否自动将星期几的第一个字母大写。
    /// </summary>
    bool CorrectDays { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Word 是否自动将正确的字体应用于朝鲜文文本中间键入的拉丁文单词，反之亦然。
    /// </summary>
    bool CorrectHangulAndAlphabet { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示如果单词的前两个字母以大写形式键入，Microsoft Word 是否自动将第二个字母设为小写。
    /// </summary>
    bool CorrectInitialCaps { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示如果你使用当前键盘语言以外的语言键入文本，Microsoft Word 是否自动将字词转置为其本机字母。
    /// </summary>
    bool CorrectKeyboardSetting { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Word 是否自动将每个句子中的第一个字母大写。
    /// </summary>
    bool CorrectSentenceCaps { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否自动大写表格单元格的第一个字母。
    /// </summary>
    bool CorrectTableCells { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否显示“自动更正选项”按钮。
    /// </summary>
    bool DisplayAutoCorrectOptions { get; set; }

    /// <summary>
    /// 获取表示自动更正条目当前列表的集合。
    /// </summary>
    IWordAutoCorrectEntries? Entries { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示 Microsoft Word 是否自动将缩写添加到自动更正首字母例外列表。
    /// </summary>
    bool FirstLetterAutoAdd { get; set; }

    /// <summary>
    /// 获取表示 Microsoft Word 不会自动大写下一个字母的缩写列表的集合。
    /// </summary>
    IWordFirstLetterExceptions? FirstLetterExceptions { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否自动将单词添加到“自动更正例外项”对话框中“朝鲜文”选项卡上的朝鲜文和字母自动更正异常列表。
    /// </summary>
    bool HangulAndAlphabetAutoAdd { get; set; }

    /// <summary>
    /// 获取表示朝鲜文和字母自动更正异常列表的集合。
    /// </summary>
    IWordHangulAndAlphabetExceptions? HangulAndAlphabetExceptions { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否向“自动更正异常”对话框中“其他更正”选项卡上的“自动更正”异常列表添加字词。
    /// </summary>
    bool OtherCorrectionsAutoAdd { get; set; }

    /// <summary>
    /// 获取表示 Microsoft Word 不会自动更正的字词列表的集合。
    /// </summary>
    IWordOtherCorrectionsExceptions? OtherCorrectionsExceptions { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否自动将指定的文本替换为“自动更正”列表中的条目。
    /// </summary>
    bool ReplaceText { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在用户键入时自动将拼写错误的文本替换为拼写检查器的建议。
    /// </summary>
    bool ReplaceTextFromSpellingChecker { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否自动将单词添加到自动更正初始大写异常列表。
    /// </summary>
    bool TwoInitialCapsAutoAdd { get; set; }

    /// <summary>
    /// 获取表示 Microsoft Word 不会自动更正的包含混合大写的术语列表的集合。
    /// </summary>
    MsWord.TwoInitialCapsExceptions TwoInitialCapsExceptions { get; }

    #endregion
}