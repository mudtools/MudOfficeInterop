//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 语言ID枚举
/// 用于指定文档或文本的语言设置
/// </summary>
public enum WdLanguageID
{
    /// <summary>
    /// 无语言指定
    /// </summary>
    wdLanguageNone = 0,

    /// <summary>
    /// 中国香港特别行政区中文
    /// </summary>
    wdChineseHongKongSAR = 3076,

    /// <summary>
    /// 中国澳门特别行政区中文
    /// </summary>
    wdChineseMacaoSAR = 5124,

    /// <summary>
    /// 简体中文
    /// </summary>
    wdSimplifiedChinese = 2052,

    /// <summary>
    /// 新加坡中文
    /// </summary>
    wdChineseSingapore = 4100,

    /// <summary>
    /// 繁体中文
    /// </summary>
    wdTraditionalChinese = 1028,

    /// <summary>
    /// 澳大利亚英语
    /// </summary>
    wdEnglishAUS = 3081,

    /// <summary>
    /// 英国英语
    /// </summary>
    wdEnglishUK = 2057,

    /// <summary>
    /// 美国英语
    /// </summary>
    wdEnglishUS = 1033,

    /// <summary>
    /// 日语
    /// </summary>
    wdJapanese = 1041,

    /// <summary>
    /// 韩语
    /// </summary>
    wdKorean = 1042,

    /// <summary>
    /// 法语
    /// </summary>
    wdFrench = 1036,

    /// <summary>
    /// 德语
    /// </summary>
    wdGerman = 1031,

    /// <summary>
    /// 西班牙语
    /// </summary>
    wdSpanish = 1034,
}