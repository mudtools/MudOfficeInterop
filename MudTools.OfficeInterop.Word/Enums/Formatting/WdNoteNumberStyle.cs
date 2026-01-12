//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定脚注和尾注的编号样式
/// </summary>
public enum WdNoteNumberStyle
{
    /// <summary>
    /// 阿拉伯数字（1, 2, 3...）
    /// </summary>
    wdNoteNumberStyleArabic = 0,
    
    /// <summary>
    /// 大写罗马数字（I, II, III...）
    /// </summary>
    wdNoteNumberStyleUppercaseRoman = 1,
    
    /// <summary>
    /// 小写罗马数字（i, ii, iii...）
    /// </summary>
    wdNoteNumberStyleLowercaseRoman = 2,
    
    /// <summary>
    /// 大写英文字母（A, B, C...）
    /// </summary>
    wdNoteNumberStyleUppercaseLetter = 3,
    
    /// <summary>
    /// 小写英文字母（a, b, c...）
    /// </summary>
    wdNoteNumberStyleLowercaseLetter = 4,
    
    /// <summary>
    /// 符号
    /// </summary>
    wdNoteNumberStyleSymbol = 9,
    
    /// <summary>
    /// 全角阿拉伯数字
    /// </summary>
    wdNoteNumberStyleArabicFullWidth = 14,
    
    /// <summary>
    /// 汉字数字（KANJI）
    /// </summary>
    wdNoteNumberStyleKanji = 10,
    
    /// <summary>
    /// 汉字数位（KANJI DIGIT）
    /// </summary>
    wdNoteNumberStyleKanjiDigit = 11,
    
    /// <summary>
    /// 传统汉字数字
    /// </summary>
    wdNoteNumberStyleKanjiTraditional = 16,
    
    /// <summary>
    /// 圆圈中的数字
    /// </summary>
    wdNoteNumberStyleNumberInCircle = 18,
    
    /// <summary>
    /// 韩文字母读音
    /// </summary>
    wdNoteNumberStyleHanjaRead = 41,
    
    /// <summary>
    /// 韩文字母数位读音
    /// </summary>
    wdNoteNumberStyleHanjaReadDigit = 42,
    
    /// <summary>
    /// 传统中文数字1
    /// </summary>
    wdNoteNumberStyleTradChinNum1 = 33,
    
    /// <summary>
    /// 传统中文数字2
    /// </summary>
    wdNoteNumberStyleTradChinNum2 = 34,
    
    /// <summary>
    /// 简体中文数字1
    /// </summary>
    wdNoteNumberStyleSimpChinNum1 = 37,
    
    /// <summary>
    /// 简体中文数字2
    /// </summary>
    wdNoteNumberStyleSimpChinNum2 = 38,
    
    /// <summary>
    /// 希伯来字母1
    /// </summary>
    wdNoteNumberStyleHebrewLetter1 = 45,
    
    /// <summary>
    /// 阿拉伯字母1
    /// </summary>
    wdNoteNumberStyleArabicLetter1 = 46,
    
    /// <summary>
    /// 希伯来字母2
    /// </summary>
    wdNoteNumberStyleHebrewLetter2 = 47,
    
    /// <summary>
    /// 阿拉伯字母2
    /// </summary>
    wdNoteNumberStyleArabicLetter2 = 48,
    
    /// <summary>
    /// 印地语字母1
    /// </summary>
    wdNoteNumberStyleHindiLetter1 = 49,
    
    /// <summary>
    /// 印地语字母2
    /// </summary>
    wdNoteNumberStyleHindiLetter2 = 50,
    
    /// <summary>
    /// 印地语阿拉伯数字
    /// </summary>
    wdNoteNumberStyleHindiArabic = 51,
    
    /// <summary>
    /// 印地语基数文本
    /// </summary>
    wdNoteNumberStyleHindiCardinalText = 52,
    
    /// <summary>
    /// 泰语字母
    /// </summary>
    wdNoteNumberStyleThaiLetter = 53,
    
    /// <summary>
    /// 泰语阿拉伯数字
    /// </summary>
    wdNoteNumberStyleThaiArabic = 54,
    
    /// <summary>
    /// 泰语基数文本
    /// </summary>
    wdNoteNumberStyleThaiCardinalText = 55,
    
    /// <summary>
    /// 越南语基数文本
    /// </summary>
    wdNoteNumberStyleVietCardinalText = 56
}