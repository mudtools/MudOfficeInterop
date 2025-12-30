//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示 Microsoft Word 中用于电子邮件的全局首选项。
/// <para>注：使用 Application.EmailOptions 属性可返回 EmailOptions 对象。</para>
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordEmailOptions : IDisposable
{
    #region 基本属性 (Basic Properties)

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
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    #endregion

    /// <summary>
    /// 获取或设置一个值，指示新的电子邮件是否使用默认电子邮件主题定义的字符样式。
    /// </summary>
    bool UseThemeStyle { get; set; }

    /// <summary>
    /// 获取或设置 Microsoft Word 在电子邮件中标记注释时使用的字符串。
    /// </summary>
    string MarkCommentsWith { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否在电子邮件中标记用户的注释。
    /// </summary>
    bool MarkComments { get; set; }

    /// <summary>
    /// 获取 EmailSignature 对象，表示 Microsoft Word 附加到传出电子邮件的签名。
    /// </summary>
    IWordEmailSignature? EmailSignature { get; }

    /// <summary>
    /// 获取表示用于撰写新电子邮件样式的 Style 对象。
    /// </summary>
    IWordStyle? ComposeStyle { get; }

    /// <summary>
    /// 获取表示回复电子邮件时使用的样式的 Style 对象。
    /// </summary>
    IWordStyle? ReplyStyle { get; }

    /// <summary>
    /// 获取或设置用于新电子邮件的主题名称及任何主题格式化选项。
    /// </summary>
    string ThemeName { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示回复电子邮件时是否为回复文本选择新颜色。
    /// </summary>
    bool NewColorOnReply { get; set; }

    /// <summary>
    /// 获取表示纯文本格式电子邮件文本属性的 Style 对象。
    /// </summary>
    IWordStyle? PlainTextStyle { get; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 在回复电子邮件时是否使用主题。
    /// </summary>
    bool UseThemeStyleOnReply { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入时是否自动将样式应用于标题。
    /// </summary>
    bool AutoFormatAsYouTypeApplyHeadings { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示按下 ENTER 键时是否将三个或更多连字符（-）、等号（=）或下划线字符（_）自动替换为特定边框线。
    /// </summary>
    bool AutoFormatAsYouTypeApplyBorders { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入时是否将项目符号字符（如星号、连字符和大于号）替换为“项目符号和编号”对话框中的项目符号。
    /// </summary>
    bool AutoFormatAsYouTypeApplyBulletedLists { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否根据键入的内容自动将段落格式化为编号列表。
    /// </summary>
    bool AutoFormatAsYouTypeApplyNumberedLists { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入时是否自动将直引号更改为智能（弯曲）引号。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceQuotes { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入时是否将两个连续的连字符（--）替换为短破折号（–）或长破折号（—）。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceSymbols { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入时是否将序数词后缀“st”、“nd”、“rd”和“th”替换为上标字母。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceOrdinals { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入时是否将键入的分数替换为当前字符集中的分数。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceFractions { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入时是否将手动强调字符自动替换为字符格式。
    /// </summary>
    bool AutoFormatAsYouTypeReplacePlainTextEmphasis { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否将应用于列表项开头的字符格式重复到下一个列表项。
    /// </summary>
    bool AutoFormatAsYouTypeFormatListItemBeginning { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否根据手动格式自动创建新样式。
    /// </summary>
    bool AutoFormatAsYouTypeDefineStyles { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入时是否自动将电子邮件地址、服务器和共享名称（UNC 路径）以及 Internet 地址（URL）更改为超链接。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceHyperlinks { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入加号、一系列连字符、另一个加号等然后按 ENTER 时，Microsoft Word 是否自动创建表格。
    /// </summary>
    bool AutoFormatAsYouTypeApplyTables { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否自动将段落开头输入的空格替换为首行缩进。
    /// </summary>
    bool AutoFormatAsYouTypeApplyFirstIndents { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入时是否自动将日期样式应用于日期。
    /// </summary>
    bool AutoFormatAsYouTypeApplyDates { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入时是否自动将结束样式应用于信件结尾。
    /// </summary>
    bool AutoFormatAsYouTypeApplyClosings { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否自动更正不匹配的括号。
    /// </summary>
    bool AutoFormatAsYouTypeMatchParentheses { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否自动更正长元音和破折号。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceFarEastDashes { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示键入时是否自动删除在日文和拉丁文本之间插入的空格。
    /// </summary>
    bool AutoFormatAsYouTypeDeleteAutoSpaces { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示输入备忘录标题时，Microsoft Word 是否自动插入相应的备忘录结尾。
    /// </summary>
    bool AutoFormatAsYouTypeInsertClosings { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示输入信件问候语或结尾时，Microsoft Word 是否自动启动“信件向导”。
    /// </summary>
    bool AutoFormatAsYouTypeAutoLetterWizard { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示输入“”或“”时，Microsoft Word 是否自动插入“ ”。
    /// </summary>
    bool AutoFormatAsYouTypeInsertOvers { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在 Web 浏览器中查看保存的文档时是否使用级联样式表（CSS）进行字体格式化。
    /// </summary>
    bool RelyOnCSS { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否去除用于在 Microsoft Word 中打开 HTML 文件但显示不需要的 HTML 标签。
    /// </summary>
    WdEmailHTMLFidelity HTMLFidelity { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否可以使用 TAB 键增加段落左缩进，使用 BACKSPACE 键减少左缩进，以及是否可以使用 BACKSPACE 键将右对齐段落更改为居中对齐，将居中对齐段落更改为左对齐。
    /// </summary>
    bool TabIndentKey { get; set; }
}