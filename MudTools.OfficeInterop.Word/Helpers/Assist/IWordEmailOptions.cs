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
public interface IWordEmailOptions : IDisposable
{
    #region 基本属性 (Basic Properties)

    /// <summary>
    /// 获取与该对象关联的 Word 应用程序。
    /// </summary>
    IWordApplication Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    /// </summary>
    int Creator { get; }

    #endregion

    #region 电子邮件选项属性 (Email Options Properties)

    /// <summary>
    /// 获取或设置一个值，该值指示是否在新电子邮件中使用主题行作为邮件标题。
    /// </summary>
    bool UseThemeStyle { get; set; }

    /// <summary>
    /// 获取或设置用于新电子邮件的主题行样式。
    /// </summary>
    string ThemeName { get; set; }

    /// <summary>
    /// 获取或设置用于新电子邮件的纯文本字体。
    /// </summary>
    IWordStyle PlainTextStyle { get; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在新电子邮件中使用 Microsoft Outlook 的主题样式。
    /// </summary>
    bool RelyOnCSS { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在新电子邮件中使用级联样式表 (CSS)。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceQuotes { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动将直引号替换为弯引号。
    /// </summary>
    bool AutoFormatAsYouTypeApplyBorders { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动应用段落边框。
    /// </summary>
    bool AutoFormatAsYouTypeApplyBulletedLists { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动应用项目符号列表。
    /// </summary>
    bool AutoFormatAsYouTypeApplyNumberedLists { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动应用编号列表。
    /// </summary>
    bool AutoFormatAsYouTypeApplyHeadings { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动应用标题样式。
    /// </summary>
    bool AutoFormatAsYouTypeFormatListItemBeginning { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动设置列表项开头的格式。
    /// </summary>
    bool AutoFormatAsYouTypeDefineStyles { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动定义新样式。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceSymbols { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动将分数 (如 1/2) 替换为符号。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceOrdinals { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动将序数词 (如 1st) 的后缀设置为上标。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceFractions { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动将分数 (如 1/2) 设置为分数格式。
    /// </summary>
    bool AutoFormatAsYouTypeReplacePlainTextEmphasis { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动将纯文本强调格式（如 *bold*）替换为真正的格式。
    /// </summary>
    bool AutoFormatAsYouTypeReplaceHyperlinks { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动将电子邮件地址和网址设置为超链接。
    /// </summary>
    bool AutoFormatAsYouTypeApplyTables { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动创建表格。
    /// </summary>
    bool AutoFormatAsYouTypeApplyFirstIndents { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在键入时自动应用首行缩进。
    /// </summary>
    bool AutoFormatAsYouTypeApplyDates { get; set; }

    /// <summary>
    /// 获取或设置一个值，该值指示是否在滚动时实时更新屏幕显示。
    /// </summary>
    bool MarkComments { get; set; }
    #endregion
}