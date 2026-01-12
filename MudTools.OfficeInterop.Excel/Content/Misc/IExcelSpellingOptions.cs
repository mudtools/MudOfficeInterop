//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示工作表的多种拼写检查选项。
/// </summary>
[ComObjectWrap(ComNamespace = "MsExcel")]
public interface IExcelSpellingOptions : IOfficeObject<IExcelSpellingOptions, MsExcel.SpellingOptions>, IDisposable
{
    /// <summary>
    /// 获取或设置 Microsoft Excel 执行拼写检查时使用的词典语言。
    /// </summary>
    int DictLang { get; set; }

    /// <summary>
    /// 获取或设置自定义词典的路径，在对工作表执行拼写检查时，新单词可以添加到该词典中。
    /// </summary>
    string UserDict { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在使用拼写检查器时是否忽略大写单词。False 表示检查大写单词；True 表示忽略大写单词。
    /// </summary>
    bool IgnoreCaps { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在使用拼写检查器时是否仅从主词典中建议单词。True 表示仅从主词典中建议单词；False 表示取消此限制。
    /// </summary>
    bool SuggestMainOnly { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示检查拼写时是否忽略混合数字。False 表示检查混合数字；True 表示忽略混合数字。
    /// </summary>
    bool IgnoreMixedDigits { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在使用拼写检查器时是否忽略 Internet 和文件地址。False 表示检查 Internet 和文件地址；True 表示忽略 Internet 和文件地址。
    /// </summary>
    bool IgnoreFileNames { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示是否使用德语后改革规则检查单词的拼写。True 表示使用德语后改革规则；False 表示取消此功能。
    /// </summary>
    bool GermanPostReform { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在检查拼写时是否合并韩语助动词和形容词。
    /// </summary>
    bool KoreanCombineAux { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在使用拼写检查器时是否启用韩语单词的自动更改列表。
    /// </summary>
    bool KoreanUseAutoChangeList { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示在使用拼写检查器时是否处理韩语复合名词。
    /// </summary>
    bool KoreanProcessCompound { get; set; }


    /// <summary>
    /// 获取或设置一个值，指示拼写检查器是否使用关于以 alef hamza 开头的阿拉伯语单词的规则。
    /// </summary>
    bool ArabicStrictAlefHamza { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示拼写检查器是否使用关于以字母 yaa 结尾的阿拉伯语单词的规则。
    /// </summary>
    bool ArabicStrictFinalYaa { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示拼写检查器是否使用规则来标记以 haa 而不是 taa marboota 结尾的阿拉伯语单词。
    /// </summary>
    bool ArabicStrictTaaMarboota { get; set; }

    /// <summary>
    /// 获取或设置一个值，指示拼写检查器是否使用关于包含字符 ë 的俄语单词的规则。
    /// </summary>
    bool RussianStrictE { get; set; }
}