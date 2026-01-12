//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定文档中页码的格式样式
/// </summary>
public enum WdPageNumberStyle
{
    /// <summary>阿拉伯数字 (1, 2, 3, ...)</summary>
    wdPageNumberStyleArabic = 0,
    /// <summary>大写罗马数字 (I, II, III, ...)</summary>
    wdPageNumberStyleUppercaseRoman = 1,
    /// <summary>小写罗马数字 (i, ii, iii, ...)</summary>
    wdPageNumberStyleLowercaseRoman = 2,
    /// <summary>大写英文字母 (A, B, C, ...)</summary>
    wdPageNumberStyleUppercaseLetter = 3,
    /// <summary>小写英文字母 (a, b, c, ...)</summary>
    wdPageNumberStyleLowercaseLetter = 4,
    /// <summary>阿拉伯数字全角 (１, ２, ３, ...)</summary>
    wdPageNumberStyleArabicFullWidth = 14,
    /// <summary>汉字数字 (〇, 一, 二, 三, ...)</summary>
    wdPageNumberStyleKanji = 10,
    /// <summary>日文汉字数字 (1, 2, 3, ...)</summary>
    wdPageNumberStyleKanjiDigit = 11,
    /// <summary>传统汉字数字 (壹, 贰, 叁, ...)</summary>
    wdPageNumberStyleKanjiTraditional = 16,
    /// <summary>带圈数字 (①, ②, ③, ...)</summary>
    wdPageNumberStyleNumberInCircle = 18,
    /// <summary>韩文汉字读音</summary>
    wdPageNumberStyleHanjaRead = 41,
    /// <summary>韩文汉字数字</summary>
    wdPageNumberStyleHanjaReadDigit = 42,
    /// <summary>繁体中文数字一 (壹, 贰, 叁, ...)</summary>
    wdPageNumberStyleTradChinNum1 = 33,
    /// <summary>繁体中文数字二 (一, 二, 三, ...)</summary>
    wdPageNumberStyleTradChinNum2 = 34,
    /// <summary>简体中文数字一 (一, 二, 三, ...)</summary>
    wdPageNumberStyleSimpChinNum1 = 37,
    /// <summary>简体中文数字二 (壹, 贰, 叁, ...)</summary>
    wdPageNumberStyleSimpChinNum2 = 38,
    /// <summary>希伯来文字符一 (א, ב, ג, ...)</summary>
    wdPageNumberStyleHebrewLetter1 = 45,
    /// <summary>阿拉伯文字符一</summary>
    wdPageNumberStyleArabicLetter1 = 46,
    /// <summary>希伯来文字符二</summary>
    wdPageNumberStyleHebrewLetter2 = 47,
    /// <summary>阿拉伯文字符二</summary>
    wdPageNumberStyleArabicLetter2 = 48,
    /// <summary>印地语字符一</summary>
    wdPageNumberStyleHindiLetter1 = 49,
    /// <summary>印地语字符二</summary>
    wdPageNumberStyleHindiLetter2 = 50,
    /// <summary>印地语阿拉伯数字</summary>
    wdPageNumberStyleHindiArabic = 51,
    /// <summary>印地语基数文本</summary>
    wdPageNumberStyleHindiCardinalText = 52,
    /// <summary>泰语字符</summary>
    wdPageNumberStyleThaiLetter = 53,
    /// <summary>泰语阿拉伯数字</summary>
    wdPageNumberStyleThaiArabic = 54,
    /// <summary>泰语基数文本</summary>
    wdPageNumberStyleThaiCardinalText = 55,
    /// <summary>越南语基数文本</summary>
    wdPageNumberStyleVietCardinalText = 56,
    /// <summary>带短横线数字 (‑1‑, ‑2‑, ‑3‑, ...)</summary>
    wdPageNumberStyleNumberInDash = 57
}