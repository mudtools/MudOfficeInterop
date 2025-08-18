//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定语言ID的枚举，用于Office应用程序的多语言支持
/// </summary>
public enum MsoLanguageID
{
    /// <summary>
    /// 混合语言ID
    /// </summary>
    msoLanguageIDMixed = -2,

    /// <summary>
    /// 无语言指定
    /// </summary>
    msoLanguageIDNone = 0,

    /// <summary>
    /// 无校对 - 1024
    /// </summary>
    msoLanguageIDNoProofing = 1024,

    /// <summary>
    /// 中国香港特别行政区中文 - 3076
    /// </summary>
    msoLanguageIDChineseHongKongSAR = 3076,

    /// <summary>
    /// 中国澳门特别行政区中文 - 5124
    /// </summary>
    msoLanguageIDChineseMacaoSAR = 5124,

    /// <summary>
    /// 简体中文 - 2052
    /// </summary>
    msoLanguageIDSimplifiedChinese = 2052,

    /// <summary>
    /// 新加坡中文 - 4100
    /// </summary>
    msoLanguageIDChineseSingapore = 4100,

    /// <summary>
    /// 繁体中文 - 1028
    /// </summary>
    msoLanguageIDTraditionalChinese = 1028,

    /// <summary>
    /// 英式英语 - 2057
    /// </summary>
    msoLanguageIDEnglishUK = 2057,

    /// <summary>
    /// 美式英语 - 1033
    /// </summary>
    msoLanguageIDEnglishUS = 1033,
}