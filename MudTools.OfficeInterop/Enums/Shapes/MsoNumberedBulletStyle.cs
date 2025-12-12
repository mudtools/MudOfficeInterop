//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定在Office应用程序中使用的编号项目符号样式
/// </summary>
public enum MsoNumberedBulletStyle
{
    /// <summary>
    /// 混合编号样式（用于表示多种样式的组合）
    /// </summary>
    msoBulletStyleMixed = -2,

    /// <summary>
    /// 小写英文字母加句点（a., b., c.）
    /// </summary>
    msoBulletAlphaLCPeriod = 0,

    /// <summary>
    /// 大写英文字母加句点（A., B., C.）
    /// </summary>
    msoBulletAlphaUCPeriod = 1,

    /// <summary>
    /// 阿拉伯数字加右括号（1), 2), 3)）
    /// </summary>
    msoBulletArabicParenRight = 2,

    /// <summary>
    /// 阿拉伯数字加句点（1., 2., 3.）
    /// </summary>
    msoBulletArabicPeriod = 3,

    /// <summary>
    /// 小写罗马数字加左右括号（(i), (ii), (iii)）
    /// </summary>
    msoBulletRomanLCParenBoth = 4,

    /// <summary>
    /// 小写罗马数字加右括号（i), ii), iii)）
    /// </summary>
    msoBulletRomanLCParenRight = 5,

    /// <summary>
    /// 小写罗马数字加句点（i., ii., iii.）
    /// </summary>
    msoBulletRomanLCPeriod = 6,

    /// <summary>
    /// 大写罗马数字加句点（I., II., III.）
    /// </summary>
    msoBulletRomanUCPeriod = 7,

    /// <summary>
    /// 小写英文字母加左右括号（(a), (b), (c)）
    /// </summary>
    msoBulletAlphaLCParenBoth = 8,

    /// <summary>
    /// 小写英文字母加右括号（a), b), c)）
    /// </summary>
    msoBulletAlphaLCParenRight = 9,

    /// <summary>
    /// 大写英文字母加左右括号（(A), (B), (C)）
    /// </summary>
    msoBulletAlphaUCParenBoth = 10,

    /// <summary>
    /// 大写英文字母加右括号（A), B), C)）
    /// </summary>
    msoBulletAlphaUCParenRight = 11,

    /// <summary>
    /// 阿拉伯数字加左右括号（(1), (2), (3)）
    /// </summary>
    msoBulletArabicParenBoth = 12,

    /// <summary>
    /// 纯阿拉伯数字（1, 2, 3）
    /// </summary>
    msoBulletArabicPlain = 13,

    /// <summary>
    /// 大写罗马数字加左右括号（(I), (II), (III)）
    /// </summary>
    msoBulletRomanUCParenBoth = 14,

    /// <summary>
    /// 大写罗马数字加右括号（I), II), III)）
    /// </summary>
    msoBulletRomanUCParenRight = 15,

    /// <summary>
    /// 简体中文数字（一, 二, 三）
    /// </summary>
    msoBulletSimpChinPlain = 16,

    /// <summary>
    /// 简体中文数字加句点（一., 二., 三.）
    /// </summary>
    msoBulletSimpChinPeriod = 17,

    /// <summary>
    /// 圆圈内的双字节数字（①, ②, ③）
    /// </summary>
    msoBulletCircleNumDBPlain = 18,

    /// <summary>
    /// 白色圆圈内的双字节数字（⓪, ①, ②）
    /// </summary>
    msoBulletCircleNumWDWhitePlain = 19,

    /// <summary>
    /// 黑色圆圈内的双字节数字（●, ●, ●）
    /// </summary>
    msoBulletCircleNumWDBlackPlain = 20,

    /// <summary>
    /// 繁体中文数字（壹, 贰, 叁）
    /// </summary>
    msoBulletTradChinPlain = 21,

    /// <summary>
    /// 繁体中文数字加句点（壹., 贰., 叁.）
    /// </summary>
    msoBulletTradChinPeriod = 22,

    /// <summary>
    /// 阿拉伯数字与字母组合加破折号（1-, 2-, 3-）
    /// </summary>
    msoBulletArabicAlphaDash = 23,

    /// <summary>
    /// 阿拉伯文字加破折号
    /// </summary>
    msoBulletArabicAbjadDash = 24,

    /// <summary>
    /// 希伯来文字加破折号
    /// </summary>
    msoBulletHebrewAlphaDash = 25,

    /// <summary>
    /// 日文汉字韩文纯数字
    /// </summary>
    msoBulletKanjiKoreanPlain = 26,

    /// <summary>
    /// 日文汉字韩文数字加句点
    /// </summary>
    msoBulletKanjiKoreanPeriod = 27,

    /// <summary>
    /// 双字节阿拉伯数字（１, ２, ３）
    /// </summary>
    msoBulletArabicDBPlain = 28,

    /// <summary>
    /// 双字节阿拉伯数字加句点（１., ２., ３.）
    /// </summary>
    msoBulletArabicDBPeriod = 29,

    /// <summary>
    /// 泰文字母加句点
    /// </summary>
    msoBulletThaiAlphaPeriod = 30,

    /// <summary>
    /// 泰文字母加右括号
    /// </summary>
    msoBulletThaiAlphaParenRight = 31,

    /// <summary>
    /// 泰文字母加左右括号
    /// </summary>
    msoBulletThaiAlphaParenBoth = 32,

    /// <summary>
    /// 泰文数字加句点
    /// </summary>
    msoBulletThaiNumPeriod = 33,

    /// <summary>
    /// 泰文数字加右括号
    /// </summary>
    msoBulletThaiNumParenRight = 34,

    /// <summary>
    /// 泰文数字加左右括号
    /// </summary>
    msoBulletThaiNumParenBoth = 35,

    /// <summary>
    /// 印地语字母加句点
    /// </summary>
    msoBulletHindiAlphaPeriod = 36,

    /// <summary>
    /// 印地语数字加句点
    /// </summary>
    msoBulletHindiNumPeriod = 37,

    /// <summary>
    /// 日文汉字简体中文双字节句点
    /// </summary>
    msoBulletKanjiSimpChinDBPeriod = 38,

    /// <summary>
    /// 印地语数字加右括号
    /// </summary>
    msoBulletHindiNumParenRight = 39,

    /// <summary>
    /// 印地语字母变体加句点
    /// </summary>
    msoBulletHindiAlpha1Period = 40
}