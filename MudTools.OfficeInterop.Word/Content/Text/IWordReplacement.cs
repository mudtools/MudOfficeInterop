
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Replacement 的接口，用于操作查找替换格式。
/// </summary>
public interface IWordReplacement : IDisposable
{
    /// <summary>
    /// 获取或设置替换文本。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置替换的字体格式。
    /// </summary>
    IWordFont? Font { get; }

    /// <summary>
    /// 获取或设置替换的段落格式。
    /// </summary>
    IWordParagraphFormat? ParagraphFormat { get; }

    /// <summary>
    /// 获取或设置替换文本的框架格式。
    /// </summary>
    IWordFrame? Frame { get; }

    /// <summary>
    /// 获取或设置替换文本的语言ID。
    /// </summary>
    WdLanguageID LanguageID { get; set; }

    /// <summary>
    /// 获取或设置替换文本的内置样式。
    /// </summary>
    WdBuiltinStyle Style { get; set; }

    /// <summary>
    /// 获取或设置替换文本的语言校对选项。值为-1时禁用拼写和语法检查，0时启用默认语言校对，其他值对应特定语言ID。
    /// </summary>
    int NoProofing { get; set; }

    /// <summary>
    /// 获取或设置替换的行距。
    /// </summary>
    float LineSpacing { get; set; }

    /// <summary>
    /// 获取或设置替换的行距规则。
    /// </summary>
    WdLineSpacing LineSpacingRule { get; set; }

    /// <summary>
    /// 获取或设置替换的段前间距。
    /// </summary>
    float SpaceBefore { get; set; }

    /// <summary>
    /// 获取或设置替换的段后间距。
    /// </summary>
    float SpaceAfter { get; set; }

    /// <summary>
    /// 获取或设置替换的首行缩进。
    /// </summary>
    float FirstLineIndent { get; set; }

    /// <summary>
    /// 获取或设置替换的左缩进。
    /// </summary>
    float LeftIndent { get; set; }

    /// <summary>
    /// 获取或设置替换的右缩进。
    /// </summary>
    float RightIndent { get; set; }

    /// <summary>
    /// 获取或设置替换的字符间距。
    /// </summary>
    float CharacterSpacing { get; set; }

    /// <summary>
    /// 获取或设置替换的字符缩放比例。
    /// </summary>
    int CharacterScaling { get; set; }

    /// <summary>
    /// 获取或设置替换的字符位置偏移。
    /// </summary>
    int Position { get; set; }

    /// <summary>
    /// 获取或设置替换的字体大小。
    /// </summary>
    float FontSize { get; set; }

    /// <summary>
    /// 获取或设置替换的字体名称。
    /// </summary>
    string FontName { get; set; }

    /// <summary>
    /// 获取或设置替换的粗体状态。
    /// </summary>
    bool Bold { get; set; }

    /// <summary>
    /// 获取或设置替换的斜体状态。
    /// </summary>
    bool Italic { get; set; }

    /// <summary>
    /// 获取或设置替换的下划线状态。
    /// </summary>
    bool Underline { get; set; }

    /// <summary>
    /// 获取或设置替换的上标状态。
    /// </summary>
    bool Superscript { get; set; }

    /// <summary>
    /// 获取或设置替换的下标状态。
    /// </summary>
    bool Subscript { get; set; }

    /// <summary>
    /// 获取或设置替换的突出显示状态。
    /// </summary>
    int Highlight { get; set; }

    /// <summary>
    /// 清除所有替换格式。
    /// </summary>
    void ClearFormatting();
}