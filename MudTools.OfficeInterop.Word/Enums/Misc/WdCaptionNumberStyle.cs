//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定与 CaptionLabel 对象一起使用的题注编号样式
/// </summary>
public enum WdCaptionNumberStyle
{
    /// <summary>
    /// 阿拉伯数字样式
    /// </summary>
    wdCaptionNumberStyleArabic = 0,

    /// <summary>
    /// 大写罗马数字样式
    /// </summary>
    wdCaptionNumberStyleUppercaseRoman = 1,

    /// <summary>
    /// 小写罗马数字样式
    /// </summary>
    wdCaptionNumberStyleLowercaseRoman = 2,

    /// <summary>
    /// 大写字母样式
    /// </summary>
    wdCaptionNumberStyleUppercaseLetter = 3,

    /// <summary>
    /// 小写字母样式
    /// </summary>
    wdCaptionNumberStyleLowercaseLetter = 4,

    /// <summary>
    /// 全角阿拉伯数字样式
    /// </summary>
    wdCaptionNumberStyleArabicFullWidth = 14,

    /// <summary>
    /// 日文汉字样式
    /// </summary>
    wdCaptionNumberStyleKanji = 10,

    /// <summary>
    /// 日文汉字数字样式
    /// </summary>
    wdCaptionNumberStyleKanjiDigit = 11,

    /// <summary>
    /// 日文传统样式
    /// </summary>
    wdCaptionNumberStyleKanjiTraditional = 16,

    /// <summary>
    /// 圆圈数字样式
    /// </summary>
    wdCaptionNumberStyleNumberInCircle = 18,

    /// <summary>
    /// 韩文ganada样式
    /// </summary>
    wdCaptionNumberStyleGanada = 24,

    /// <summary>
    /// 韩文chosung样式
    /// </summary>
    wdCaptionNumberStyleChosung = 25,

    /// <summary>
    /// 中文天干地支样式1
    /// </summary>
    wdCaptionNumberStyleZodiac1 = 30,

    /// <summary>
    /// 中文天干地支样式2
    /// </summary>
    wdCaptionNumberStyleZodiac2 = 31,

    /// <summary>
    /// 韩文汉字读法样式
    /// </summary>
    wdCaptionNumberStyleHanjaRead = 41,

    /// <summary>
    /// 韩文汉字数字读法样式
    /// </summary>
    wdCaptionNumberStyleHanjaReadDigit = 42,

    /// <summary>
    /// 繁体中文数字样式2
    /// </summary>
    wdCaptionNumberStyleTradChinNum2 = 34,

    /// <summary>
    /// 繁体中文数字样式3
    /// </summary>
    wdCaptionNumberStyleTradChinNum3 = 35,

    /// <summary>
    /// 简体中文数字样式2
    /// </summary>
    wdCaptionNumberStyleSimpChinNum2 = 38,

    /// <summary>
    /// 简体中文数字样式3
    /// </summary>
    wdCaptionNumberStyleSimpChinNum3 = 39,

    /// <summary>
    /// 希伯来字母样式1
    /// </summary>
    wdCaptionNumberStyleHebrewLetter1 = 45,

    /// <summary>
    /// 阿拉伯字母样式1
    /// </summary>
    wdCaptionNumberStyleArabicLetter1 = 46,

    /// <summary>
    /// 希伯来字母样式2
    /// </summary>
    wdCaptionNumberStyleHebrewLetter2 = 47,

    /// <summary>
    /// 阿拉伯字母样式2
    /// </summary>
    wdCaptionNumberStyleArabicLetter2 = 48,

    /// <summary>
    /// 印地文字母样式1
    /// </summary>
    wdCaptionNumberStyleHindiLetter1 = 49,

    /// <summary>
    /// 印地文字母样式2
    /// </summary>
    wdCaptionNumberStyleHindiLetter2 = 50,

    /// <summary>
    /// 印地文阿拉伯数字样式
    /// </summary>
    wdCaptionNumberStyleHindiArabic = 51,

    /// <summary>
    /// 印地文基数词样式
    /// </summary>
    wdCaptionNumberStyleHindiCardinalText = 52,

    /// <summary>
    /// 泰文字母样式
    /// </summary>
    wdCaptionNumberStyleThaiLetter = 53,

    /// <summary>
    /// 泰文阿拉伯数字样式
    /// </summary>
    wdCaptionNumberStyleThaiArabic = 54,

    /// <summary>
    /// 泰文基数词样式
    /// </summary>
    wdCaptionNumberStyleThaiCardinalText = 55,

    /// <summary>
    /// 越南文基数词样式
    /// </summary>
    wdCaptionNumberStyleVietCardinalText = 56
}