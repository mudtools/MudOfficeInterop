//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定渲染文本时要使用的字符集
/// </summary>
public enum MsoCharacterSet
{
    /// <summary>
    /// 阿拉伯字符集
    /// </summary>
    msoCharacterSetArabic = 1,

    /// <summary>
    /// 西里尔字符集
    /// </summary>
    msoCharacterSetCyrillic,

    /// <summary>
    /// 英语、西欧及其他拉丁文字符集
    /// </summary>
    msoCharacterSetEnglishWesternEuropeanOtherLatinScript,

    /// <summary>
    /// 希腊字符集
    /// </summary>
    msoCharacterSetGreek,

    /// <summary>
    /// 希伯来字符集
    /// </summary>
    msoCharacterSetHebrew,

    /// <summary>
    /// 日文字符集
    /// </summary>
    msoCharacterSetJapanese,

    /// <summary>
    /// 韩文字符集
    /// </summary>
    msoCharacterSetKorean,

    /// <summary>
    /// 多语言Unicode字符集
    /// </summary>
    msoCharacterSetMultilingualUnicode,

    /// <summary>
    /// 简体中文字符集
    /// </summary>
    msoCharacterSetSimplifiedChinese,

    /// <summary>
    /// 泰文字符集
    /// </summary>
    msoCharacterSetThai,

    /// <summary>
    /// 繁体中文字符集
    /// </summary>
    msoCharacterSetTraditionalChinese,

    /// <summary>
    /// 越南文字符集
    /// </summary>
    msoCharacterSetVietnamese
}