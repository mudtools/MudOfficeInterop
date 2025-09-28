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
public interface IWordFind : IDisposable
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
    /// 获取或设置查找文本
    /// </summary>
    string FindText { get; set; }

    /// <summary>
    /// 获取或设置替换文本
    /// </summary>
    string ReplaceWith { get; set; }

    /// <summary>
    /// 获取或设置是否区分大小写
    /// </summary>
    bool MatchCase { get; set; }

    /// <summary>
    /// 获取或设置是否全字匹配
    /// </summary>
    bool MatchWholeWord { get; set; }

    /// <summary>
    /// 获取或设置是否使用通配符
    /// </summary>
    bool MatchWildcards { get; set; }

    /// <summary>
    /// 获取或设置是否匹配前缀
    /// </summary>
    bool MatchPrefix { get; set; }

    /// <summary>
    /// 获取或设置是否匹配后缀
    /// </summary>
    bool MatchSuffix { get; set; }

    /// <summary>
    /// 获取或设置是否忽略空格
    /// </summary>
    bool IgnoreSpace { get; set; }

    /// <summary>
    /// 获取或设置是否忽略标点符号
    /// </summary>
    bool IgnorePunct { get; set; }

    /// <summary>
    /// 获取或设置查找包装方式
    /// </summary>
    WdFindWrap Wrap { get; set; }

    /// <summary>
    /// 获取或设置是否匹配格式
    /// </summary>
    bool Format { get; set; }

    /// <summary>
    /// 获取或设置不校对选项
    /// </summary>
    int NoProofing { get; set; }

    /// <summary>
    /// 获取或设置是否匹配控制字符
    /// </summary>
    bool MatchControl { get; set; }

    /// <summary>
    /// 获取或设置是否匹配短语
    /// </summary>
    bool MatchPhrase { get; set; }

    /// <summary>
    /// 获取或设置是否匹配发音相似的内容
    /// </summary>
    bool MatchSoundsLike { get; set; }

    /// <summary>
    /// 获取或设置是否匹配所有词形变化
    /// </summary>
    bool MatchAllWordForms { get; set; }

    /// <summary>
    /// 获取或设置是否按字节匹配
    /// </summary>
    bool MatchByte { get; set; }

    /// <summary>
    /// 获取或设置是否模糊匹配
    /// </summary>
    bool MatchFuzzy { get; set; }

    /// <summary>
    /// 获取是否找到匹配项
    /// </summary>
    bool Found { get; }

    /// <summary>
    /// 获取框架对象
    /// </summary>
    IWordFrame? Frame { get; }

    /// <summary>
    /// 获取字体对象
    /// </summary>
    IWordFont Font { get; }

    /// <summary>
    /// 获取段落格式对象
    /// </summary>
    IWordParagraphFormat ParagraphFormat { get; }

    /// <summary>
    /// 获取替换对象
    /// </summary>
    IWordReplacement Replacement { get; }

    /// <summary>
    /// 获取或设置是否高亮显示
    /// </summary>
    bool Highlight { get; set; }

    /// <summary>
    /// 获取或设置语言标识符
    /// </summary>
    WdLanguageID LanguageID { get; set; }

    /// <summary>
    /// 获取或设置样式
    /// </summary>
    WdBuiltinStyle Style { get; set; }

    /// <summary>
    /// 获取或设置是否向前查找
    /// </summary>
    bool Forward { get; set; }

    /// <summary>
    /// 执行查找或替换操作
    /// </summary>
    /// <param name="findText">要查找的文本内容</param>
    /// <param name="matchCase">是否区分大小写，默认为false</param>
    /// <param name="matchWholeWord">是否全字匹配，默认为false</param>
    /// <param name="matchWildcards">是否使用通配符，默认为false</param>
    /// <param name="matchSoundsLike">是否匹配发音相似的词，默认为null</param>
    /// <param name="matchAllWordForms">是否匹配所有词形变化，默认为null</param>
    /// <param name="forward">查找方向，true表示向前查找，false表示向后查找，默认为null</param>
    /// <param name="wrap">查找包装方式，默认为null</param>
    /// <param name="format">是否匹配格式，默认为null</param>
    /// <param name="replaceWith">替换文本内容，默认为null</param>
    /// <param name="replace">替换选项，默认为null</param>
    /// <returns>如果成功找到匹配项则返回true，否则返回false</returns>
    bool Execute(string? findText = null, bool? matchCase = false,
        bool? matchWholeWord = false, bool? matchWildcards = false,
        bool? matchSoundsLike = null, bool? matchAllWordForms = null,
        bool? forward = null, WdFindWrap? wrap = null, bool? format = null,
        string? replaceWith = null, WdReplace? replace = null);

    /// <summary>
    /// 突出显示找到的文本内容
    /// </summary>
    /// <param name="findText">要查找的文本内容</param>
    /// <param name="highlightColor">高亮颜色，默认为黄色</param>
    /// <param name="textColor">文本颜色，默认为黑色</param>
    /// <param name="matchCase">是否区分大小写，默认为false</param>
    /// <param name="matchWholeWord">是否全字匹配</param>
    /// <param name="matchPrefix">是否匹配前缀</param>
    /// <param name="matchSuffix">是否匹配后缀</param>
    /// <param name="matchPhrase">是否匹配短语</param>
    /// <param name="matchWildcards">是否使用通配符</param>
    /// <param name="matchSoundsLike">是否匹配发音相似的词</param>
    /// <param name="matchAllWordForms">是否匹配所有词形变化</param>
    /// <param name="matchByte">是否按字节匹配</param>
    /// <param name="matchFuzzy">是否模糊匹配</param>
    /// <param name="ignoreSpace">是否忽略空格</param>
    /// <param name="ignorePunct">是否忽略标点符号</param>
    /// <returns>是否成功执行高亮操作</returns>
    bool HitHighlight(string? findText = null,
          WdColor? highlightColor = WdColor.wdColorYellow, WdColor? textColor = WdColor.wdColorBlack,
          bool? matchCase = false, bool? matchWholeWord = null,
          bool? matchPrefix = null, bool? matchSuffix = null,
          bool? matchPhrase = null, bool? matchWildcards = null,
          bool? matchSoundsLike = null, bool? matchAllWordForms = null,
          bool? matchByte = null, bool? matchFuzzy = null,
          bool? ignoreSpace = null, bool? ignorePunct = null);

    /// <summary>
    /// 执行查找并替换
    /// </summary>
    /// <param name="replace">替换选项</param>
    /// <returns>是否找到并替换</returns>
    bool ExecuteReplace(WdReplace replace = WdReplace.wdReplaceAll);

    /// <summary>
    /// 清除查找设置
    /// </summary>
    void ClearFormatting();

    /// <summary>
    /// 清除替换设置
    /// </summary>
    void ClearReplaceFormatting();

    /// <summary>
    /// 清除高亮显示设置
    /// </summary>
    void ClearHitHighlight();
}
