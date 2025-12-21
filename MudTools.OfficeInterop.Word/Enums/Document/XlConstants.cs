//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定 Microsoft Word 中的各种常量
/// </summary>
public enum XlConstants
{
    /// <summary>
    /// Microsoft Word 对指定对象应用自动设置，如颜色或页码
    /// </summary>
    xlAutomatic = -4105,

    /// <summary>
    /// 组合图
    /// </summary>
    xlCombination = -4111,

    /// <summary>
    /// Microsoft Word 对指定对象应用自定义设置，如颜色或误差量
    /// </summary>
    xlCustom = -4114,

    /// <summary>
    /// 二维条形图组或系列
    /// </summary>
    xlBar = 2,

    /// <summary>
    /// 柱形图组或系列
    /// </summary>
    xlColumn = 3,

    /// <summary>
    /// 三维条形图组或系列
    /// </summary>
    xl3DBar = -4099,

    /// <summary>
    /// 三维曲面图组或系列
    /// </summary>
    xl3DSurface = -4103,

    /// <summary>
    /// Microsoft Word 应用默认或自动格式
    /// </summary>
    xlDefaultAutoFormat = -1,

    /// <summary>
    /// 不在指定图表组或系列中显示误差线
    /// </summary>
    xlNone = -4142,

    /// <summary>
    /// 汇总行显示在指定范围上方
    /// </summary>
    xlAbove = 0,

    /// <summary>
    /// 汇总行显示在指定范围下方
    /// </summary>
    xlBelow = 1,

    /// <summary>
    /// 在指定图表组或系列中同时显示正负误差线
    /// </summary>
    xlBoth = 1,

    /// <summary>
    /// 底部
    /// </summary>
    xlBottom = -4107,

    /// <summary>
    /// 居中
    /// </summary>
    xlCenter = -4108,

    /// <summary>
    /// 棋盘图案
    /// </summary>
    xlChecker = 9,

    /// <summary>
    /// 圆形
    /// </summary>
    xlCircle = 8,

    /// <summary>
    /// 角点
    /// </summary>
    xlCorner = 2,

    /// <summary>
    /// 十字交叉图案
    /// </summary>
    xlCrissCross = 16,

    /// <summary>
    /// 交叉图案
    /// </summary>
    xlCross = 4,

    /// <summary>
    /// 菱形图案
    /// </summary>
    xlDiamond = 2,

    /// <summary>
    /// 分散对齐
    /// </summary>
    xlDistributed = -4117,

    /// <summary>
    /// 填充
    /// </summary>
    xlFill = 5,

    /// <summary>
    /// 将误差量显示为固定值
    /// </summary>
    xlFixedValue = 1,

    /// <summary>
    /// 常规
    /// </summary>
    xlGeneral = 1,

    /// <summary>
    /// 16%灰色图案
    /// </summary>
    xlGray16 = 17,

    /// <summary>
    /// 25%灰色图案
    /// </summary>
    xlGray25 = -4124,

    /// <summary>
    /// 50%灰色图案
    /// </summary>
    xlGray50 = -4125,

    /// <summary>
    /// 75%灰色图案
    /// </summary>
    xlGray75 = -4126,

    /// <summary>
    /// 8%灰色图案
    /// </summary>
    xlGray8 = 18,

    /// <summary>
    /// 网格图案
    /// </summary>
    xlGrid = 15,

    /// <summary>
    /// 高
    /// </summary>
    xlHigh = -4127,

    /// <summary>
    /// 内部
    /// </summary>
    xlInside = 2,

    /// <summary>
    /// 两端对齐
    /// </summary>
    xlJustify = -4130,

    /// <summary>
    /// 左对齐
    /// </summary>
    xlLeft = -4131,

    /// <summary>
    /// 浅色向下线条图案
    /// </summary>
    xlLightDown = 13,

    /// <summary>
    /// 浅色水平线条图案
    /// </summary>
    xlLightHorizontal = 11,

    /// <summary>
    /// 浅色向上线条图案
    /// </summary>
    xlLightUp = 14,

    /// <summary>
    /// 浅色垂直线条图案
    /// </summary>
    xlLightVertical = 12,

    /// <summary>
    /// 低
    /// </summary>
    xlLow = -4134,

    /// <summary>
    /// 最大值
    /// </summary>
    xlMaximum = 2,

    /// <summary>
    /// 最小值
    /// </summary>
    xlMinimum = 4,

    /// <summary>
    /// 负值
    /// </summary>
    xlMinusValues = 3,

    /// <summary>
    /// 靠近坐标轴
    /// </summary>
    xlNextToAxis = 4,

    /// <summary>
    /// 不透明填充
    /// </summary>
    xlOpaque = 3,

    /// <summary>
    /// 外部
    /// </summary>
    xlOutside = 3,

    /// <summary>
    /// 将误差量显示为百分比
    /// </summary>
    xlPercent = 2,

    /// <summary>
    /// 在指定图表组或系列中显示正误差线
    /// </summary>
    xlPlus = 9,

    /// <summary>
    /// 正值
    /// </summary>
    xlPlusValues = 2,

    /// <summary>
    /// 右对齐
    /// </summary>
    xlRight = -4152,

    /// <summary>
    /// 缩放
    /// </summary>
    xlScale = 3,

    /// <summary>
    /// 75%半灰色图案
    /// </summary>
    xlSemiGray75 = 10,

    /// <summary>
    /// 显示标签
    /// </summary>
    xlShowLabel = 4,

    /// <summary>
    /// 显示标签和百分比
    /// </summary>
    xlShowLabelAndPercent = 5,

    /// <summary>
    /// 显示百分比
    /// </summary>
    xlShowPercent = 3,

    /// <summary>
    /// 显示数值
    /// </summary>
    xlShowValue = 2,

    /// <summary>
    /// 单线
    /// </summary>
    xlSingle = 2,

    /// <summary>
    /// 实心图案
    /// </summary>
    xlSolid = 1,

    /// <summary>
    /// 正方形
    /// </summary>
    xlSquare = 1,

    /// <summary>
    /// 星形
    /// </summary>
    xlStar = 5,

    /// <summary>
    /// 将误差量显示为标准误差
    /// </summary>
    xlStError = 4,

    /// <summary>
    /// 顶部
    /// </summary>
    xlTop = -4160,

    /// <summary>
    /// 透明填充
    /// </summary>
    xlTransparent = 2,

    /// <summary>
    /// 三角形
    /// </summary>
    xlTriangle = 3
}