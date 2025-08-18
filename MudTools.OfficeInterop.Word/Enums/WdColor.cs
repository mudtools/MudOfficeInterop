//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// Word 颜色枚举
/// 包含 Word 中使用的标准颜色值
/// </summary>
public enum WdColor
{
    /// <summary>
/// 自动颜色（根据应用程序决定实际颜色）
/// </summary>
wdColorAutomatic = -16777216,
    
    /// <summary>
    /// 黑色
    /// </summary>
    wdColorBlack = 0,
    
    /// <summary>
    /// 蓝色
    /// </summary>
    wdColorBlue = 16711680,
    
    /// <summary>
    /// 青绿色
    /// </summary>
    wdColorTurquoise = 16776960,
    
    /// <summary>
    /// 亮绿色
    /// </summary>
    wdColorBrightGreen = 65280,
    
    /// <summary>
    /// 粉色
    /// </summary>
    wdColorPink = 16711935,
    
    /// <summary>
    /// 红色
    /// </summary>
    wdColorRed = 255,
    
    /// <summary>
    /// 黄色
    /// </summary>
    wdColorYellow = 65535,
    
    /// <summary>
    /// 白色
    /// </summary>
    wdColorWhite = 16777215,
    
    /// <summary>
    /// 深蓝色
    /// </summary>
    wdColorDarkBlue = 8388608,
    
    /// <summary>
    /// 蓝绿色
    /// </summary>
    wdColorTeal = 8421376,
    
    /// <summary>
    /// 绿色
    /// </summary>
    wdColorGreen = 32768,
    
    /// <summary>
    /// 紫色
    /// </summary>
    wdColorViolet = 8388736,
    
    /// <summary>
    /// 深红色
    /// </summary>
    wdColorDarkRed = 128,
    
    /// <summary>
    /// 深黄色
    /// </summary>
    wdColorDarkYellow = 32896,
    
    /// <summary>
    /// 棕色
    /// </summary>
    wdColorBrown = 13209,
    
    /// <summary>
    /// 橄榄绿
    /// </summary>
    wdColorOliveGreen = 13107,
    
    /// <summary>
    /// 深绿色
    /// </summary>
    wdColorDarkGreen = 13056,
    
    /// <summary>
    /// 深蓝绿色
    /// </summary>
    wdColorDarkTeal = 6697728,
    
    /// <summary>
    /// 靛蓝色
    /// </summary>
    wdColorIndigo = 10040115,
    
    /// <summary>
    /// 橙色
    /// </summary>
    wdColorOrange = 26367,
    
    /// <summary>
    /// 蓝灰色
    /// </summary>
    wdColorBlueGray = 10053222,
    
    /// <summary>
    /// 浅橙色
    /// </summary>
    wdColorLightOrange = 39423,
    
    /// <summary>
    /// 酸橙色
    /// </summary>
    wdColorLime = 52377,
    
    /// <summary>
    /// 海绿色
    /// </summary>
    wdColorSeaGreen = 6723891,
    
    /// <summary>
    /// 水色
    /// </summary>
    wdColorAqua = 13421619,
    
    /// <summary>
    /// 浅蓝色
    /// </summary>
    wdColorLightBlue = 16737843,
    
    /// <summary>
    /// 金色
    /// </summary>
    wdColorGold = 52479,
    
    /// <summary>
    /// 天蓝色
    /// </summary>
    wdColorSkyBlue = 16763904,
    
    /// <summary>
    /// 梅红色
    /// </summary>
    wdColorPlum = 6697881,
    
    /// <summary>
    /// 玫瑰色
    /// </summary>
    wdColorRose = 13408767,
    
    /// <summary>
    /// 黄褐色
    /// </summary>
    wdColorTan = 10079487,
    
    /// <summary>
    /// 浅黄色
    /// </summary>
    wdColorLightYellow = 10092543,
    
    /// <summary>
    /// 浅绿色
    /// </summary>
    wdColorLightGreen = 13434828,
    
    /// <summary>
    /// 浅青绿色
    /// </summary>
    wdColorLightTurquoise = 16777164,
    
    /// <summary>
    /// 淡蓝色
    /// </summary>
    wdColorPaleBlue = 16764057,
    
    /// <summary>
    /// 薰衣草色
    /// </summary>
    wdColorLavender = 16751052,
    
    /// <summary>
    /// 5% 灰色
    /// </summary>
    wdColorGray05 = 15987699,
    
    /// <summary>
    /// 10% 灰色
    /// </summary>
    wdColorGray10 = 15132390,
    
    /// <summary>
    /// 12.5% 灰色
    /// </summary>
    wdColorGray125 = 14737632,
    
    /// <summary>
    /// 15% 灰色
    /// </summary>
    wdColorGray15 = 14277081,
    
    /// <summary>
    /// 20% 灰色
    /// </summary>
    wdColorGray20 = 13421772,
    
    /// <summary>
    /// 25% 灰色
    /// </summary>
    wdColorGray25 = 12632256,
    
    /// <summary>
    /// 30% 灰色
    /// </summary>
    wdColorGray30 = 11776947,
    
    /// <summary>
    /// 35% 灰色
    /// </summary>
    wdColorGray35 = 10921638,
    
    /// <summary>
    /// 37.5% 灰色
    /// </summary>
    wdColorGray375 = 10526880,
    
    /// <summary>
    /// 40% 灰色
    /// </summary>
    wdColorGray40 = 10066329,
    
    /// <summary>
    /// 45% 灰色
    /// </summary>
    wdColorGray45 = 9211020,
    
    /// <summary>
    /// 50% 灰色
    /// </summary>
    wdColorGray50 = 8421504,
    
    /// <summary>
    /// 55% 灰色
    /// </summary>
    wdColorGray55 = 7566195,
    
    /// <summary>
    /// 60% 灰色
    /// </summary>
    wdColorGray60 = 6710886,
    
    /// <summary>
    /// 62.5% 灰色
    /// </summary>
    wdColorGray625 = 6316128,
    
    /// <summary>
    /// 65% 灰色
    /// </summary>
    wdColorGray65 = 5855577,
    
    /// <summary>
    /// 70% 灰色
    /// </summary>
    wdColorGray70 = 5000268,
    
    /// <summary>
    /// 75% 灰色
    /// </summary>
    wdColorGray75 = 4210752,
    
    /// <summary>
    /// 80% 灰色
    /// </summary>
    wdColorGray80 = 3355443,
    
    /// <summary>
    /// 85% 灰色
    /// </summary>
    wdColorGray85 = 2500134,
    
    /// <summary>
    /// 87.5% 灰色
    /// </summary>
    wdColorGray875 = 2105376,
    
    /// <summary>
    /// 90% 灰色
    /// </summary>
    wdColorGray90 = 1644825,
    
    /// <summary>
    /// 95% 灰色
    /// </summary>
    wdColorGray95 = 789516
}