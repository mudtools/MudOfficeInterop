//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;


/// <summary>
/// Word 查找接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordFind : IDisposable
{
    /// <summary>
    /// 获取应用程序对象。
    /// </summary>
    [ComPropertyWrap(NeedDispose = false, NeedConvert = true)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取父对象。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置一个值，指示查找操作是否向前搜索文档。
    /// </summary>
    bool Forward { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示查找操作是否区分大小写。
    /// </summary>
    bool MatchCase { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示查找操作是否只定位整个单词，而不是较大单词的一部分。
    /// </summary>
    bool MatchWholeWord { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示要查找的文本是否包含通配符。
    /// </summary>
    bool MatchWildcards { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示查找操作是否返回与要查找的文本发音相似的单词。
    /// </summary>
    bool MatchSoundsLike { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示查找操作是否找到要查找的文本的所有形式（例如，如果要查找"sit"，则"sat"和"sitting"也会被找到）。
    /// </summary>
    bool MatchAllWordForms { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 在搜索期间是否为日语文本使用非特定搜索选项。
    /// </summary>
    bool MatchFuzzy { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 在搜索期间是否区分全角和半角字母或字符。
    /// </summary>
    bool MatchByte { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 是否查找或替换拼写和语法检查器忽略的文本。
    /// </summary>
    int NoProofing { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示查找操作是否在阿拉伯语文档中匹配具有匹配 kashidas 的文本。
    /// </summary>
    bool MatchKashida { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示查找操作是否在从右到左语言文档中匹配具有匹配变音符号的文本。
    /// </summary>
    bool MatchDiacritics { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示查找操作是否在阿拉伯语文档中匹配具有匹配 alef hamzas 的文本。
    /// </summary>
    bool MatchAlefHamza { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示查找操作是否在从右到左语言文档中匹配具有匹配双向控制字符的文本。
    /// </summary>
    bool MatchControl { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，指示是否忽略单词之间的所有空白和控制字符。
    /// </summary>
    bool MatchPhrase { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，指示是否匹配以搜索字符串开头的单词。
    /// </summary>
    bool MatchPrefix { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，指示是否匹配以搜索字符串结尾的单词。
    /// </summary>
    bool MatchSuffix { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，指示查找操作是否应忽略在找到的文本中的额外空白。
    /// </summary>
    bool IgnoreSpace { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，指示查找操作是否应忽略在找到的文本中的标点符号。
    /// </summary>
    bool IgnorePunct { get; set; }

    /// <summary>
    /// 获取或设置一个布尔值，指示是否在朝鲜语查找操作中定位语音韩文和汉字字符。
    /// </summary>
    bool HanjaPhoneticHangul { get; set; }

    /// <summary>
    /// 获取一个值，指示对指定对象的搜索是否产生了匹配。
    /// </summary>
    bool Found { get; }

    /// <summary>
    /// 获取或设置是否在查找条件中包含突出显示格式。
    /// </summary>
    int Highlight { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否在查找操作中包含格式。
    /// </summary>
    bool Format { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示 Microsoft Word 在用朝鲜语替换朝鲜语文本时是否自动更正朝鲜语结尾。
    /// </summary>
    bool CorrectHangulEndings { get; set; }

    /// <summary>
    /// 返回表示指定对象的字符格式的 Font 对象。
    /// </summary>
    IWordFont? Font { get; set; }

    /// <summary>
    /// 返回表示指定范围、选择、查找或替换操作或段落的段落设置的 ParagraphFormat 对象。
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
    WdBuiltinStyle? StyleType { get; set; }

    /// <summary>
    /// 获取或设置指定对象的样式。
    /// </summary>
    [ComPropertyWrap(IsMethod = true, PropertyName = "Style", NeedConvert = true)]
    string? StyleName { get; set; }

    /// <summary>
    /// 返回或设置在指定范围或选择中查找或替换的文本。
    /// </summary>
    string Text { get; set; }

    /// <summary>
    /// 返回表示替换操作条件的 Replacement 对象。
    /// </summary>
    IWordReplacement? Replacement { get; }

    /// <summary>
    /// 返回表示指定样式或查找和替换操作的框架格式的 Frame 对象。
    /// </summary>
    IWordFrame? Frame { get; }

    /// <summary>
    /// 返回或设置指定对象的语言。
    /// </summary>
    WdLanguageID LanguageID { get; set; }

    /// <summary>
    /// 返回或设置指定对象的东亚语言。
    /// </summary>
    WdLanguageID LanguageIDFarEast { get; set; }

    /// <summary>
    /// 返回或设置指定对象的语言。
    /// </summary>
    WdLanguageID LanguageIDOther { get; set; }

    /// <summary>
    /// 返回或设置如果搜索从文档中除开头以外的点开始并且到达文档末尾（或者如果 Forward 设置为 False，则反之亦然），或者如果搜索文本在指定选择或范围中未找到，会发生什么。
    /// </summary>
    WdFindWrap Wrap { get; set; }

    /// <summary>
    /// 从选择或查找或替换操作中的格式中删除文本和段落格式。
    /// </summary>
    void ClearFormatting();

    /// <summary>
    /// 激活与日语文本关联的所有非特定搜索选项。
    /// </summary>
    void SetAllFuzzyOptions();

    /// <summary>
    /// 清除与日语文本关联的所有非特定搜索选项。
    /// </summary>
    void ClearAllFuzzyOptions();

    /// <summary>
    /// 运行指定的查找操作。
    /// </summary>
    /// <param name="findText">可选 Object。要搜索的文本。使用空字符串 ("") 仅搜索格式。可以通过指定适当的字符代码来搜索特殊字符。例如，"^p" 对应于段落标记，"^t" 对应于制表符。</param>
    /// <param name="matchCase">可选 Object。如果为 True，则指定查找文本区分大小写。对应于"查找和替换"对话框（编辑菜单）中的"区分大小写"复选框。</param>
    /// <param name="matchWholeWord">可选 Object。如果为 True，则查找操作只定位整个单词，而不是较大单词的一部分。对应于"查找和替换"对话框中的"全字匹配"复选框。</param>
    /// <param name="matchWildcards">可选 Object。如果为 True，则查找文本是特殊搜索运算符。对应于"查找和替换"对话框中的"使用通配符"复选框。</param>
    /// <param name="matchSoundsLike">可选 Object。如果为 True，则查找操作定位与查找文本发音相似的单词。对应于"查找和替换"对话框中的"同音"复选框。</param>
    /// <param name="matchAllWordForms">可选 Object。如果为 True，则查找操作定位查找文本的所有形式（例如，"sit" 定位"sitting"和"sat"）。对应于"查找和替换"对话框中的"查找单词的所有形式"复选框。</param>
    /// <param name="forward">可选 Object。如果为 True，则向前搜索（朝向文档末尾）。</param>
    /// <param name="wrap">可选 Object。控制在搜索从文档中除开头以外的点开始并且到达文档末尾（或者如果 Forward 设置为 False，则反之亦然）时会发生什么。此参数还控制如果有选择或范围并且搜索文本在选择或范围中未找到时会发生什么。可以是以下 WdFindWrap 常量之一：wdFindAsk 在搜索选择或范围后，Microsoft Word 显示消息询问是否搜索文档的其余部分。wdFindContinue 如果到达搜索范围的开头或末尾，查找操作继续。wdFindStop 如果到达搜索范围的开头或末尾，查找操作结束。</param>
    /// <param name="format">可选 Object。如果为 True，则查找操作定位格式，除了或代替查找文本。</param>
    /// <param name="replaceWith">可选 Object。替换文本。要删除 Find 参数指定的文本，请使用空字符串 ("")。可以像为 Find 参数一样指定特殊字符和高级搜索条件。要将图形对象或其他非文本项目指定为替换，请将项目移动到剪贴板并为 ReplaceWith 指定"^c"。</param>
    /// <param name="replace">可选 Object。指定要进行的替换数量：一个、全部或没有。可以是任何 WdReplace 常量：wdReplaceAll、wdReplaceNone、wdReplaceOne。</param>
    /// <param name="matchKashida">可选 Object。如果为 True，则查找操作在阿拉伯语文档中匹配具有匹配 kashidas 的文本。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="matchDiacritics">可选 Object。如果为 True，则查找操作在从右到左语言文档中匹配具有匹配变音符号的文本。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="matchAlefHamza">可选 Object。如果为 True，则查找操作在阿拉伯语文档中匹配具有匹配 alef hamzas 的文本。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="matchControl">可选 Object。如果为 True，则查找操作在从右到左语言文档中匹配具有匹配双向控制字符的文本。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <returns>如果查找操作成功，则为 True；否则为 False。</returns>
    bool? Execute(string? findText = null, bool? matchCase = null, bool? matchWholeWord = null, bool? matchWildcards = null,
                bool? matchSoundsLike = null, bool? matchAllWordForms = null, bool? forward = null,
                WdFindWrap? wrap = null, bool? format = null, string? replaceWith = null, WdReplace? replace = null,
                bool? matchKashida = null, bool? matchDiacritics = null, bool? matchAlefHamza = null,
                bool? matchControl = null);

    /// <summary>
    /// 高亮显示所有找到的匹配项，并返回一个布尔值，表示是否找到匹配项。
    /// </summary>
    /// <param name="findText">指定要查找的文本。使用空字符串 ("") 仅搜索格式。可以通过指定适当的字符代码来搜索特殊字符。例如，"^p" 对应于段落标记，"^t" 对应于制表符。</param>
    /// <param name="highlightColor">指定文本的高亮颜色。可以是任何 RGB 颜色或 WdColor 枚举值之一。</param>
    /// <param name="textColor">指定文本的颜色。可以是任何 RGB 颜色或 WdColor 枚举值之一。</param>
    /// <param name="matchCase">如果为 True，则指定查找文本区分大小写。对应于"查找和替换"对话框中的"区分大小写"复选框。</param>
    /// <param name="matchWholeWord">如果为 True，则查找操作只定位整个单词，而不是较大单词的一部分。对应于"查找和替换"对话框中的"全字匹配"复选框。</param>
    /// <param name="matchPrefix">如果为 True，则匹配以搜索字符串开头的单词。对应于"查找和替换"对话框中的"匹配前缀"复选框。</param>
    /// <param name="matchSuffix">如果为 True，则匹配以搜索字符串结尾的单词。对应于"查找和替换"对话框中的"匹配后缀"复选框。</param>
    /// <param name="matchPhrase">如果为 True，则忽略单词之间的所有空白和控制字符。</param>
    /// <param name="matchWildcards">如果为 True，则查找文本是特殊搜索运算符。对应于"查找和替换"对话框中的"使用通配符"复选框。</param>
    /// <param name="matchSoundsLike">如果为 True，则查找操作定位与查找文本发音相似的单词。对应于"查找和替换"对话框中的"同音"复选框。</param>
    /// <param name="matchAllWordForms">如果为 True，则查找操作定位查找文本的所有形式（例如，"sit" 定位"sitting"和"sat"）。对应于"查找和替换"对话框中的"查找单词的所有形式"复选框。</param>
    /// <param name="matchByte">如果为 True，则在搜索期间区分全角和半角字母或字符。</param>
    /// <param name="matchFuzzy">如果为 True，则在搜索期间使用日语文本的非特定搜索选项。</param>
    /// <param name="matchKashida">如果为 True，则查找操作在阿拉伯语文档中匹配具有匹配 kashidas 的文本。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="matchDiacritics">如果为 True，则查找操作在从右到左语言文档中匹配具有匹配变音符号的文本。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="matchAlefHamza">如果为 True，则查找操作在阿拉伯语文档中匹配具有匹配 alef hamzas 的文本。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="matchControl">如果为 True，则查找操作在从右到左语言文档中匹配具有匹配双向控制字符的文本。根据您选择或安装的语言支持（例如美国英语），此参数可能不可用。</param>
    /// <param name="ignoreSpace">如果为 True，则忽略单词之间的所有空白。对应于"查找和替换"对话框中的"忽略空白字符"复选框。</param>
    /// <param name="ignorePunct">如果为 True，则忽略单词之间的所有标点字符。对应于"查找和替换"对话框中的"忽略标点"复选框。</param>
    /// <param name="hanjaPhoneticHangul">如果为 True，则忽略语音韩文和汉字字符。仅当您有朝鲜语语言支持时才可用。</param>
    /// <returns>如果找到匹配项，则为 True；否则为 False。</returns>
    bool? HitHighlight(string? findText, WdColor? highlightColor = null, WdColor? textColor = null,
                        bool? matchCase = null, bool? matchWholeWord = null, bool? matchPrefix = null,
                        bool? matchSuffix = null, bool? matchPhrase = null, bool? matchWildcards = null,
                        bool? matchSoundsLike = null, bool? matchAllWordForms = null, bool? matchByte = null,
                        bool? matchFuzzy = null, bool? matchKashida = null, bool? matchDiacritics = null,
                        bool? matchAlefHamza = null, bool? matchControl = null, bool? ignoreSpace = null,
                        bool? ignorePunct = null, bool? hanjaPhoneticHangul = null);

    /// <summary>
    /// 移除在高亮显示查找操作中定位的所有文本的高亮显示，并返回一个布尔值，表示操作是否成功。
    /// </summary>
    /// <returns>如果操作成功，则为 True；否则为 False。</returns>
    bool? ClearHitHighlight();
}
