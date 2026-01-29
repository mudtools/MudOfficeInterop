//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定编号项目符号的样式。
/// </summary>
public enum PpNumberedBulletStyle
{
    /// <summary>
    /// 混合项目符号样式。
    /// </summary>
    ppBulletStyleMixed = -2,

    /// <summary>
    /// 小写字母加句点（a., b., c.）。
    /// </summary>
    ppBulletAlphaLCPeriod = 0,

    /// <summary>
    /// 大写字母加句点（A., B., C.）。
    /// </summary>
    ppBulletAlphaUCPeriod = 1,

    /// <summary>
    /// 阿拉伯数字加右括号（1), 2), 3)）。
    /// </summary>
    ppBulletArabicParenRight = 2,

    /// <summary>
    /// 阿拉伯数字加句点（1., 2., 3.）。
    /// </summary>
    ppBulletArabicPeriod = 3,

    /// <summary>
    /// 小写罗马数字加双括号（(i), (ii), (iii)）。
    /// </summary>
    ppBulletRomanLCParenBoth = 4,

    /// <summary>
    /// 小写罗马数字加右括号（i), ii), iii)）。
    /// </summary>
    ppBulletRomanLCParenRight = 5,

    /// <summary>
    /// 小写罗马数字加句点（i., ii., iii.）。
    /// </summary>
    ppBulletRomanLCPeriod = 6,

    /// <summary>
    /// 大写罗马数字加句点（I., II., III.）。
    /// </summary>
    ppBulletRomanUCPeriod = 7,

    /// <summary>
    /// 小写字母加双括号（(a), (b), (c)）。
    /// </summary>
    ppBulletAlphaLCParenBoth = 8,

    /// <summary>
    /// 小写字母加右括号（a), b), c)）。
    /// </summary>
    ppBulletAlphaLCParenRight = 9,

    /// <summary>
    /// 大写字母加双括号（(A), (B), (C)）。
    /// </summary>
    ppBulletAlphaUCParenBoth = 10,

    /// <summary>
    /// 大写字母加右括号（A), B), C)）。
    /// </summary>
    ppBulletAlphaUCParenRight = 11,

    /// <summary>
    /// 阿拉伯数字加双括号（(1), (2), (3)）。
    /// </summary>
    ppBulletArabicParenBoth = 12,

    /// <summary>
    /// 阿拉伯数字（无标点）（1, 2, 3）。
    /// </summary>
    ppBulletArabicPlain = 13,

    /// <summary>
    /// 大写罗马数字加双括号（(I), (II), (III)）。
    /// </summary>
    ppBulletRomanUCParenBoth = 14,

    /// <summary>
    /// 大写罗马数字加右括号（I), II), III)）。
    /// </summary>
    ppBulletRomanUCParenRight = 15,

    /// <summary>
    /// 简体中文（无标点）（一、二、三）。
    /// </summary>
    ppBulletSimpChinPlain = 16,

    /// <summary>
    /// 简体中文加句点（一., 二., 三.）。
    /// </summary>
    ppBulletSimpChinPeriod = 17,

    /// <summary>
    /// 圆圈数字（双字节）无标点（①, ②, ③）。
    /// </summary>
    ppBulletCircleNumDBPlain = 18,

    /// <summary>
    /// 圆圈数字（宽字符）白底无标点（❶, ❷, ❸）。
    /// </summary>
    ppBulletCircleNumWDWhitePlain = 19,

    /// <summary>
    /// 圆圈数字（宽字符）黑底无标点（➊, ➋, ➌）。
    /// </summary>
    ppBulletCircleNumWDBlackPlain = 20,

    /// <summary>
    /// 繁体中文（无标点）（壹、貳、參）。
    /// </summary>
    ppBulletTradChinPlain = 21,

    /// <summary>
    /// 繁体中文加句点（壹., 貳., 參.）。
    /// </summary>
    ppBulletTradChinPeriod = 22,

    /// <summary>
    /// 阿拉伯字母加连字符（أ-، ب-، ج-）。
    /// </summary>
    ppBulletArabicAlphaDash = 23,

    /// <summary>
    /// 阿拉伯字母（阿布贾德）加连字符（أ-، ب-، ج-）。
    /// </summary>
    ppBulletArabicAbjadDash = 24,

    /// <summary>
    /// 希伯来字母加连字符（א-، ב-، ג-）。
    /// </summary>
    ppBulletHebrewAlphaDash = 25,

    /// <summary>
    /// 日文/韩文汉字（无标点）（一、二、三）。
    /// </summary>
    ppBulletKanjiKoreanPlain = 26,

    /// <summary>
    /// 日文/韩文汉字加句点（一., 二., 三.）。
    /// </summary>
    ppBulletKanjiKoreanPeriod = 27,

    /// <summary>
    /// 阿拉伯数字（双字节）无标点（１, ２, ３）。
    /// </summary>
    ppBulletArabicDBPlain = 28,

    /// <summary>
    /// 阿拉伯数字（双字节）加句点（１., ２., ３.）。
    /// </summary>
    ppBulletArabicDBPeriod = 29,

    /// <summary>
    /// 泰文字母加句点（ก., ข., ค.）。
    /// </summary>
    ppBulletThaiAlphaPeriod = 30,

    /// <summary>
    /// 泰文字母加右括号（ก), ข), ค)）。
    /// </summary>
    ppBulletThaiAlphaParenRight = 31,

    /// <summary>
    /// 泰文字母加双括号（(ก), (ข), (ค)）。
    /// </summary>
    ppBulletThaiAlphaParenBoth = 32,

    /// <summary>
    /// 泰文数字加句点（๑., ๒., ๓.）。
    /// </summary>
    ppBulletThaiNumPeriod = 33,

    /// <summary>
    /// 泰文数字加右括号（๑), ๒), ๓)）。
    /// </summary>
    ppBulletThaiNumParenRight = 34,

    /// <summary>
    /// 泰文数字加双括号（(๑), (๒), (๓)）。
    /// </summary>
    ppBulletThaiNumParenBoth = 35,

    /// <summary>
    /// 印地文字母加句点（क., ख., ग.）。
    /// </summary>
    ppBulletHindiAlphaPeriod = 36,

    /// <summary>
    /// 印地文数字加句点（१., २., ३.）。
    /// </summary>
    ppBulletHindiNumPeriod = 37,

    /// <summary>
    /// 日文汉字（简体中文双字节）加句点（一., 二., 三.）。
    /// </summary>
    ppBulletKanjiSimpChinDBPeriod = 38,

    /// <summary>
    /// 印地文数字加右括号（१), २), ३)）。
    /// </summary>
    ppBulletHindiNumParenRight = 39,

    /// <summary>
    /// 印地文字母第一种加句点（क., ख., ग.）。
    /// </summary>
    ppBulletHindiAlpha1Period = 40
}