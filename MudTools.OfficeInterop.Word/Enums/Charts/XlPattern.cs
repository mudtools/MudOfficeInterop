//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定应用于图表区域或填充对象的图案类型
/// </summary>
public enum XlPattern
{
    /// <summary>
    /// 自动图案
    /// </summary>
    xlPatternAutomatic = -4105,
    /// <summary>
    /// 棋盘图案
    /// </summary>
    xlPatternChecker = 9,
    /// <summary>
    /// 交叉线图案
    /// </summary>
    xlPatternCrissCross = 16,
    /// <summary>
    /// 向下斜线图案
    /// </summary>
    xlPatternDown = -4121,
    /// <summary>
    /// 16%灰度图案
    /// </summary>
    xlPatternGray16 = 17,
    /// <summary>
    /// 25%灰度图案
    /// </summary>
    xlPatternGray25 = -4124,
    /// <summary>
    /// 50%灰度图案
    /// </summary>
    xlPatternGray50 = -4125,
    /// <summary>
    /// 75%灰度图案
    /// </summary>
    xlPatternGray75 = -4126,
    /// <summary>
    /// 8%灰度图案
    /// </summary>
    xlPatternGray8 = 18,
    /// <summary>
    /// 网格图案
    /// </summary>
    xlPatternGrid = 15,
    /// <summary>
    /// 水平线图案
    /// </summary>
    xlPatternHorizontal = -4128,
    /// <summary>
    /// 浅色向下斜线图案
    /// </summary>
    xlPatternLightDown = 13,
    /// <summary>
    /// 浅色水平线图案
    /// </summary>
    xlPatternLightHorizontal = 11,
    /// <summary>
    /// 浅色向上斜线图案
    /// </summary>
    xlPatternLightUp = 14,
    /// <summary>
    /// 浅色垂直线图案
    /// </summary>
    xlPatternLightVertical = 12,
    /// <summary>
    /// 无图案
    /// </summary>
    xlPatternNone = -4142,
    /// <summary>
    /// 75%半灰度图案
    /// </summary>
    xlPatternSemiGray75 = 10,
    /// <summary>
    /// 纯色图案
    /// </summary>
    xlPatternSolid = 1,
    /// <summary>
    /// 向上斜线图案
    /// </summary>
    xlPatternUp = -4162,
    /// <summary>
    /// 垂直线图案
    /// </summary>
    xlPatternVertical = -4166,
    /// <summary>
    /// 线性渐变图案
    /// </summary>
    xlPatternLinearGradient = 4000,
    /// <summary>
    /// 矩形渐变图案
    /// </summary>
    xlPatternRectangularGradient = 4001
}