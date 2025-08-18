//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定图案类型的枚举，用于形状或对象的图案填充样式
/// </summary>
public enum MsoPatternType
{
    /// <summary>
    /// 混合图案类型
    /// </summary>
    msoPatternMixed = -2,

    /// <summary>
    /// 5%密度的图案
    /// </summary>
    msoPattern5Percent = 1,

    /// <summary>
    /// 10%密度的图案
    /// </summary>
    msoPattern10Percent = 2,

    /// <summary>
    /// 20%密度的图案
    /// </summary>
    msoPattern20Percent = 3,

    /// <summary>
    /// 25%密度的图案
    /// </summary>
    msoPattern25Percent = 4,

    /// <summary>
    /// 30%密度的图案
    /// </summary>
    msoPattern30Percent = 5,

    /// <summary>
    /// 40%密度的图案
    /// </summary>
    msoPattern40Percent = 6,

    /// <summary>
    /// 50%密度的图案
    /// </summary>
    msoPattern50Percent = 7,

    /// <summary>
    /// 60%密度的图案
    /// </summary>
    msoPattern60Percent = 8,

    /// <summary>
    /// 70%密度的图案
    /// </summary>
    msoPattern70Percent = 9,

    /// <summary>
    /// 75%密度的图案
    /// </summary>
    msoPattern75Percent = 10,

    /// <summary>
    /// 80%密度的图案
    /// </summary>
    msoPattern80Percent = 11,

    /// <summary>
    /// 90%密度的图案
    /// </summary>
    msoPattern90Percent = 12,

    /// <summary>
    /// 深色水平线图案
    /// </summary>
    msoPatternDarkHorizontal = 13,

    /// <summary>
    /// 深色垂直线图案
    /// </summary>
    msoPatternDarkVertical = 14,

    /// <summary>
    /// 深色向下对角线图案
    /// </summary>
    msoPatternDarkDownwardDiagonal = 15,

    /// <summary>
    /// 深色向上对角线图案
    /// </summary>
    msoPatternDarkUpwardDiagonal = 16,

    /// <summary>
    /// 小棋盘图案
    /// </summary>
    msoPatternSmallCheckerBoard = 17,

    /// <summary>
    /// 格子图案
    /// </summary>
    msoPatternTrellis = 18,

    /// <summary>
    /// 浅色水平线图案
    /// </summary>
    msoPatternLightHorizontal = 19,

    /// <summary>
    /// 浅色垂直线图案
    /// </summary>
    msoPatternLightVertical = 20,

    /// <summary>
    /// 浅色向下对角线图案
    /// </summary>
    msoPatternLightDownwardDiagonal = 21,

    /// <summary>
    /// 浅色向上对角线图案
    /// </summary>
    msoPatternLightUpwardDiagonal = 22,

    /// <summary>
    /// 小网格图案
    /// </summary>
    msoPatternSmallGrid = 23,

    /// <summary>
    /// 点状菱形图案
    /// </summary>
    msoPatternDottedDiamond = 24,

    /// <summary>
    /// 宽向下对角线图案
    /// </summary>
    msoPatternWideDownwardDiagonal = 25,

    /// <summary>
    /// 宽向上对角线图案
    /// </summary>
    msoPatternWideUpwardDiagonal = 26,

    /// <summary>
    /// 虚线上向上对角线图案
    /// </summary>
    msoPatternDashedUpwardDiagonal = 27,

    /// <summary>
    /// 虚线下向下对角线图案
    /// </summary>
    msoPatternDashedDownwardDiagonal = 28,

    /// <summary>
    /// 窄垂直线图案
    /// </summary>
    msoPatternNarrowVertical = 29,

    /// <summary>
    /// 窄水平线图案
    /// </summary>
    msoPatternNarrowHorizontal = 30,

    /// <summary>
    /// 虚线垂直线图案
    /// </summary>
    msoPatternDashedVertical = 31,

    /// <summary>
    /// 虚线水平线图案
    /// </summary>
    msoPatternDashedHorizontal = 32,

    /// <summary>
    /// 大五彩纸屑图案
    /// </summary>
    msoPatternLargeConfetti = 33,

    /// <summary>
    /// 大网格图案
    /// </summary>
    msoPatternLargeGrid = 34,

    /// <summary>
    /// 水平砖块图案
    /// </summary>
    msoPatternHorizontalBrick = 35,

    /// <summary>
    /// 大棋盘图案
    /// </summary>
    msoPatternLargeCheckerBoard = 36,

    /// <summary>
    /// 小五彩纸屑图案
    /// </summary>
    msoPatternSmallConfetti = 37,

    /// <summary>
    /// 之字形图案
    /// </summary>
    msoPatternZigZag = 38,

    /// <summary>
    /// 实心菱形图案
    /// </summary>
    msoPatternSolidDiamond = 39,

    /// <summary>
    /// 对角砖块图案
    /// </summary>
    msoPatternDiagonalBrick = 40,

    /// <summary>
    /// 轮廓菱形图案
    /// </summary>
    msoPatternOutlinedDiamond = 41,

    /// <summary>
    /// 格子布图案
    /// </summary>
    msoPatternPlaid = 42,

    /// <summary>
    /// 球体图案
    /// </summary>
    msoPatternSphere = 43,

    /// <summary>
    /// 编织图案
    /// </summary>
    msoPatternWeave = 44,

    /// <summary>
    /// 点状网格图案
    /// </summary>
    msoPatternDottedGrid = 45,

    /// <summary>
    /// 草皮图案
    /// </summary>
    msoPatternDivot = 46,

    /// <summary>
    /// 木瓦图案
    /// </summary>
    msoPatternShingle = 47,

    /// <summary>
    /// 波浪图案
    /// </summary>
    msoPatternWave = 48,

    /// <summary>
    /// 水平线图案
    /// </summary>
    msoPatternHorizontal = 49,

    /// <summary>
    /// 垂直线图案
    /// </summary>
    msoPatternVertical = 50,

    /// <summary>
    /// 十字图案
    /// </summary>
    msoPatternCross = 51,

    /// <summary>
    /// 向下对角线图案
    /// </summary>
    msoPatternDownwardDiagonal = 52,

    /// <summary>
    /// 向上对角线图案
    /// </summary>
    msoPatternUpwardDiagonal = 53,

    /// <summary>
    /// 对角十字图案
    /// </summary>
    msoPatternDiagonalCross = 54
}