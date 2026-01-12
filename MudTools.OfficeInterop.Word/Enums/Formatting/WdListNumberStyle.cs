//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定 Word 文档中列表项的编号样式
/// </summary>
public enum WdListNumberStyle
{
    /// <summary>
    /// 阿拉伯数字（1, 2, 3...）
    /// </summary>
    wdListNumberStyleArabic = 0,

    /// <summary>
    /// 大写罗马数字（I, II, III...）
    /// </summary>
    wdListNumberStyleUppercaseRoman = 1,

    /// <summary>
    /// 小写罗马数字（i, ii, iii...）
    /// </summary>
    wdListNumberStyleLowercaseRoman = 2,

    /// <summary>
    /// 大写英文字母（A, B, C...）
    /// </summary>
    wdListNumberStyleUppercaseLetter = 3,

    /// <summary>
    /// 小写英文字母（a, b, c...）
    /// </summary>
    wdListNumberStyleLowercaseLetter = 4,

    /// <summary>
    /// 序数数字（1st, 2nd, 3rd...）
    /// </summary>
    wdListNumberStyleOrdinal = 5,

    /// <summary>
    /// 基数文本（One, Two, Three...）
    /// </summary>
    wdListNumberStyleCardinalText = 6,

    /// <summary>
    /// 序数文本（First, Second, Third...）
    /// </summary>
    wdListNumberStyleOrdinalText = 7,

    /// <summary>
    /// 日文汉字
    /// </summary>
    wdListNumberStyleKanji = 10,

    /// <summary>
    /// 日文汉字数位
    /// </summary>
    wdListNumberStyleKanjiDigit = 11,

    /// <summary>
    /// 日文五十音图（半角）
    /// </summary>
    wdListNumberStyleAiueoHalfWidth = 12,

    /// <summary>
    /// 日文伊吕波（半角）
    /// </summary>
    wdListNumberStyleIrohaHalfWidth = 13,

    /// <summary>
    /// 全角阿拉伯数字
    /// </summary>
    wdListNumberStyleArabicFullWidth = 14,

    /// <summary>
    /// 传统日文汉字
    /// </summary>
    wdListNumberStyleKanjiTraditional = 16,

    /// <summary>
    /// 传统日文汉字2
    /// </summary>
    wdListNumberStyleKanjiTraditional2 = 17,

    /// <summary>
    /// 圆圈中的数字
    /// </summary>
    wdListNumberStyleNumberInCircle = 18,

    /// <summary>
    /// 日文五十音图
    /// </summary>
    wdListNumberStyleAiueo = 20,

    /// <summary>
    /// 日文伊吕波
    /// </summary>
    wdListNumberStyleIroha = 21,

    /// <summary>
    /// 阿拉伯数字带前导零（01, 02, 03...）
    /// </summary>
    wdListNumberStyleArabicLZ = 22,

    /// <summary>
    /// 项目符号
    /// </summary>
    wdListNumberStyleBullet = 23,

    /// <summary>
    /// 韩文谚文
    /// </summary>
    wdListNumberStyleGanada = 24,

    /// <summary>
    /// 韩文初声
    /// </summary>
    wdListNumberStyleChosung = 25,

    /// <summary>
    /// 中文数字样式1
    /// </summary>
    wdListNumberStyleGBNum1 = 26,

    /// <summary>
    /// 中文数字样式2
    /// </summary>
    wdListNumberStyleGBNum2 = 27,

    /// <summary>
    /// 中文数字样式3
    /// </summary>
    wdListNumberStyleGBNum3 = 28,

    /// <summary>
    /// 中文数字样式4
    /// </summary>
    wdListNumberStyleGBNum4 = 29,

    /// <summary>
    /// 十二生肖1
    /// </summary>
    wdListNumberStyleZodiac1 = 30,

    /// <summary>
    /// 十二生肖2
    /// </summary>
    wdListNumberStyleZodiac2 = 31,

    /// <summary>
    /// 十二生肖3
    /// </summary>
    wdListNumberStyleZodiac3 = 32,

    /// <summary>
    /// 繁体中文数字1
    /// </summary>
    wdListNumberStyleTradChinNum1 = 33,

    /// <summary>
    /// 繁体中文数字2
    /// </summary>
    wdListNumberStyleTradChinNum2 = 34,

    /// <summary>
    /// 繁体中文数字3
    /// </summary>
    wdListNumberStyleTradChinNum3 = 35,

    /// <summary>
    /// 繁体中文数字4
    /// </summary>
    wdListNumberStyleTradChinNum4 = 36,

    /// <summary>
    /// 简体中文数字1
    /// </summary>
    wdListNumberStyleSimpChinNum1 = 37,

    /// <summary>
    /// 简体中文数字2
    /// </summary>
    wdListNumberStyleSimpChinNum2 = 38,

    /// <summary>
    /// 简体中文数字3
    /// </summary>
    wdListNumberStyleSimpChinNum3 = 39,

    /// <summary>
    /// 简体中文数字4
    /// </summary>
    wdListNumberStyleSimpChinNum4 = 40,

    /// <summary>
    /// 韩文汉字读音
    /// </summary>
    wdListNumberStyleHanjaRead = 41,

    /// <summary>
    /// 韩文汉字数位读音
    /// </summary>
    wdListNumberStyleHanjaReadDigit = 42,

    /// <summary>
    /// 韩文
    /// </summary>
    wdListNumberStyleHangul = 43,

    /// <summary>
    /// 韩文汉字
    /// </summary>
    wdListNumberStyleHanja = 44,

    /// <summary>
    /// 希伯来文1
    /// </summary>
    wdListNumberStyleHebrew1 = 45,

    /// <summary>
    /// 阿拉伯文1
    /// </summary>
    wdListNumberStyleArabic1 = 46,

    /// <summary>
    /// 希伯来文2
    /// </summary>
    wdListNumberStyleHebrew2 = 47,

    /// <summary>
    /// 阿拉伯文2
    /// </summary>
    wdListNumberStyleArabic2 = 48,

    /// <summary>
    /// 印地语字母1
    /// </summary>
    wdListNumberStyleHindiLetter1 = 49,

    /// <summary>
    /// 印地语字母2
    /// </summary>
    wdListNumberStyleHindiLetter2 = 50,

    /// <summary>
    /// 印地语阿拉伯数字
    /// </summary>
    wdListNumberStyleHindiArabic = 51,

    /// <summary>
    /// 印地语基数文本
    /// </summary>
    wdListNumberStyleHindiCardinalText = 52,

    /// <summary>
    /// 泰文字母
    /// </summary>
    wdListNumberStyleThaiLetter = 53,

    /// <summary>
    /// 泰语阿拉伯数字
    /// </summary>
    wdListNumberStyleThaiArabic = 54,

    /// <summary>
    /// 泰语基数文本
    /// </summary>
    wdListNumberStyleThaiCardinalText = 55,

    /// <summary>
    /// 越南语基数文本
    /// </summary>
    wdListNumberStyleVietCardinalText = 56,

    /// <summary>
    /// 小写俄文字母
    /// </summary>
    wdListNumberStyleLowercaseRussian = 58,

    /// <summary>
    /// 大写俄文字母
    /// </summary>
    wdListNumberStyleUppercaseRussian = 59,

    /// <summary>
    /// 小写希腊字母
    /// </summary>
    wdListNumberStyleLowercaseGreek = 60,

    /// <summary>
    /// 大写希腊字母
    /// </summary>
    wdListNumberStyleUppercaseGreek = 61,

    /// <summary>
    /// 阿拉伯数字带前导零样式2
    /// </summary>
    wdListNumberStyleArabicLZ2 = 62,

    /// <summary>
    /// 阿拉伯数字带前导零样式3
    /// </summary>
    wdListNumberStyleArabicLZ3 = 63,

    /// <summary>
    /// 阿拉伯数字带前导零样式4
    /// </summary>
    wdListNumberStyleArabicLZ4 = 64,

    /// <summary>
    /// 小写土耳其字母
    /// </summary>
    wdListNumberStyleLowercaseTurkish = 65,

    /// <summary>
    /// 大写土耳其字母
    /// </summary>
    wdListNumberStyleUppercaseTurkish = 66,

    /// <summary>
    /// 小写保加利亚字母
    /// </summary>
    wdListNumberStyleLowercaseBulgarian = 67,

    /// <summary>
    /// 大写保加利亚字母
    /// </summary>
    wdListNumberStyleUppercaseBulgarian = 68,

    /// <summary>
    /// 图片项目符号
    /// </summary>
    wdListNumberStylePictureBullet = 249,

    /// <summary>
    /// 法律编号样式
    /// </summary>
    wdListNumberStyleLegal = 253,

    /// <summary>
    /// 法律编号样式带前导零
    /// </summary>
    wdListNumberStyleLegalLZ = 254,

    /// <summary>
    /// 无编号
    /// </summary>
    wdListNumberStyleNone = 255
}