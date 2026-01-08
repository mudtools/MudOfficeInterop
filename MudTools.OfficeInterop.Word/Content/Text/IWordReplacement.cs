
namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Replacement 的接口，用于操作查找替换格式。
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordReplacement : IOfficeObject<IWordReplacement, MsWord.Replacement>, IDisposable
{
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
    /// 获取或设置表示指定对象字符格式设置的Font对象。
    /// </summary>
    IWordFont? Font { get; set; }

    /// <summary>
    /// 获取或设置表示指定替换操作段落设置的ParagraphFormat对象。
    /// </summary>
    IWordParagraphFormat? ParagraphFormat { get; set; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, NeedConvert = true)]
    IWordStyle? Style { get; set; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "Style", NeedConvert = true)]
    WdBuiltinStyle StyleType { get; set; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "Style", NeedConvert = true)]
    string? StyleName { get; set; }

    /// <summary>
    /// 获取或设置在指定范围或选择中要查找或替换的文本。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 获取或设置指定对象的语言。
    /// </summary>
    WdLanguageID LanguageID { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否将突出显示格式应用于替换文本。可以返回或设置为True、False或wdUndefined。
    /// </summary>
    int Highlight { get; set; }

    /// <summary>
    /// 获取表示指定样式或查找替换操作的框架格式设置的Frame对象。
    /// </summary>
    IWordFrame? Frame { get; }

    /// <summary>
    /// 获取或设置指定对象的东亚语言。
    /// </summary>
    WdLanguageID LanguageIDFarEast { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示Microsoft Word是否查找或替换拼写和语法检查器忽略的文本。
    /// </summary>
    int NoProofing { get; set; }

    /// <summary>
    /// 从选择中或从查找或替换操作中指定的格式设置中移除文本和段落格式设置。
    /// </summary>
    void ClearFormatting();
}