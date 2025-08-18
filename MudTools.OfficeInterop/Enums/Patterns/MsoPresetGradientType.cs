//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定预设的渐变填充类型，用于Office应用程序中的形状和对象填充
/// </summary>
public enum MsoPresetGradientType
{
    /// <summary>
    /// 混合渐变类型
    /// </summary>
    msoPresetGradientMixed = -2,
    /// <summary>
    /// 早落日渐变
    /// </summary>
    msoGradientEarlySunset = 1,
    /// <summary>
    /// 晚落日渐变
    /// </summary>
    msoGradientLateSunset = 2,
    /// <summary>
    /// 夜幕降临渐变
    /// </summary>
    msoGradientNightfall = 3,
    /// <summary>
    /// 黎明渐变
    /// </summary>
    msoGradientDaybreak = 4,
    /// <summary>
    /// 地平线渐变
    /// </summary>
    msoGradientHorizon = 5,
    /// <summary>
    /// 沙漠渐变
    /// </summary>
    msoGradientDesert = 6,
    /// <summary>
    /// 海洋渐变
    /// </summary>
    msoGradientOcean = 7,
    /// <summary>
    /// 平静水面渐变
    /// </summary>
    msoGradientCalmWater = 8,
    /// <summary>
    /// 火焰渐变
    /// </summary>
    msoGradientFire = 9,
    /// <summary>
    /// 雾渐变
    /// </summary>
    msoGradientFog = 10,
    /// <summary>
    /// 苔藓渐变
    /// </summary>
    msoGradientMoss = 11,
    /// <summary>
    /// 孔雀渐变
    /// </summary>
    msoGradientPeacock = 12,
    /// <summary>
    /// 小麦渐变
    /// </summary>
    msoGradientWheat = 13,
    /// <summary>
    /// 羊皮纸渐变
    /// </summary>
    msoGradientParchment = 14,
    /// <summary>
    /// 桃花心木渐变
    /// </summary>
    msoGradientMahogany = 15,
    /// <summary>
    /// 彩虹渐变
    /// </summary>
    msoGradientRainbow = 16,
    /// <summary>
    /// 彩虹II渐变
    /// </summary>
    msoGradientRainbowII = 17,
    /// <summary>
    /// 金色渐变
    /// </summary>
    msoGradientGold = 18,
    /// <summary>
    /// 金色II渐变
    /// </summary>
    msoGradientGoldII = 19,
    /// <summary>
    /// 黄铜渐变
    /// </summary>
    msoGradientBrass = 20,
    /// <summary>
    /// 铬渐变
    /// </summary>
    msoGradientChrome = 21,
    /// <summary>
    /// 铬II渐变
    /// </summary>
    msoGradientChromeII = 22,
    /// <summary>
    /// 银色渐变
    /// </summary>
    msoGradientSilver = 23,
    /// <summary>
    /// 蓝宝石渐变
    /// </summary>
    msoGradientSapphire = 24
}