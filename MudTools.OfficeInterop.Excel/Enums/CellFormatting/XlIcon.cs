//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定图标集条件格式规则中条件的图标
/// </summary>
public enum XlIcon
{
    /// <summary>
    /// 无单元格图标
    /// </summary>
    xlIconNoCellIcon = -1,

    /// <summary>
    /// 绿色向上箭头
    /// </summary>
    xlIconGreenUpArrow = 1,

    /// <summary>
    /// 黄色侧向箭头
    /// </summary>
    xlIconYellowSideArrow = 2,

    /// <summary>
    /// 红色向下箭头
    /// </summary>
    xlIconRedDownArrow = 3,

    /// <summary>
    /// 灰色向上箭头
    /// </summary>
    xlIconGrayUpArrow = 4,

    /// <summary>
    /// 灰色侧向箭头
    /// </summary>
    xlIconGraySideArrow = 5,

    /// <summary>
    /// 灰色向下箭头
    /// </summary>
    xlIconGrayDownArrow = 6,

    /// <summary>
    /// 绿色旗帜
    /// </summary>
    xlIconGreenFlag = 7,

    /// <summary>
    /// 黄色旗帜
    /// </summary>
    xlIconYellowFlag = 8,

    /// <summary>
    /// 红色旗帜
    /// </summary>
    xlIconRedFlag = 9,

    /// <summary>
    /// 绿色圆形
    /// </summary>
    xlIconGreenCircle = 10,

    /// <summary>
    /// 黄色圆形
    /// </summary>
    xlIconYellowCircle = 11,

    /// <summary>
    /// 带边框的红色圆形
    /// </summary>
    xlIconRedCircleWithBorder = 12,

    /// <summary>
    /// 带边框的黑色圆形
    /// </summary>
    xlIconBlackCircleWithBorder = 13,

    /// <summary>
    /// 绿色交通信号灯
    /// </summary>
    xlIconGreenTrafficLight = 14,

    /// <summary>
    /// 黄色交通信号灯
    /// </summary>
    xlIconYellowTrafficLight = 15,

    /// <summary>
    /// 红色交通信号灯
    /// </summary>
    xlIconRedTrafficLight = 16,

    /// <summary>
    /// 黄色三角形
    /// </summary>
    xlIconYellowTriangle = 17,

    /// <summary>
    /// 红色菱形
    /// </summary>
    xlIconRedDiamond = 18,

    /// <summary>
    /// 绿色对勾符号
    /// </summary>
    xlIconGreenCheckSymbol = 19,

    /// <summary>
    /// 黄色感叹号符号
    /// </summary>
    xlIconYellowExclamationSymbol = 20,

    /// <summary>
    /// 红色叉号符号
    /// </summary>
    xlIconRedCrossSymbol = 21,

    /// <summary>
    /// 绿色对勾
    /// </summary>
    xlIconGreenCheck = 22,

    /// <summary>
    /// 黄色感叹号
    /// </summary>
    xlIconYellowExclamation = 23,

    /// <summary>
    /// 红色叉号
    /// </summary>
    xlIconRedCross = 24,

    /// <summary>
    /// 黄色向上斜箭头
    /// </summary>
    xlIconYellowUpInclineArrow = 25,

    /// <summary>
    /// 黄色向下斜箭头
    /// </summary>
    xlIconYellowDownInclineArrow = 26,

    /// <summary>
    /// 灰色向上斜箭头
    /// </summary>
    xlIconGrayUpInclineArrow = 27,

    /// <summary>
    /// 灰色向下斜箭头
    /// </summary>
    xlIconGrayDownInclineArrow = 28,

    /// <summary>
    /// 红色圆形
    /// </summary>
    xlIconRedCircle = 29,

    /// <summary>
    /// 粉色圆形
    /// </summary>
    xlIconPinkCircle = 30,

    /// <summary>
    /// 灰色圆形
    /// </summary>
    xlIconGrayCircle = 31,

    /// <summary>
    /// 黑色圆形
    /// </summary>
    xlIconBlackCircle = 32,

    /// <summary>
    /// 带四分之一白色的圆形
    /// </summary>
    xlIconCircleWithOneWhiteQuarter = 33,

    /// <summary>
    /// 带二分之一白色的圆形
    /// </summary>
    xlIconCircleWithTwoWhiteQuarters = 34,

    /// <summary>
    /// 带四分之三白色的圆形
    /// </summary>
    xlIconCircleWithThreeWhiteQuarters = 35,

    /// <summary>
    /// 全白色的圆形
    /// </summary>
    xlIconWhiteCircleAllWhiteQuarters = 36,

    /// <summary>
    /// 0个条形
    /// </summary>
    xlIcon0Bars = 37,

    /// <summary>
    /// 1个条形
    /// </summary>
    xlIcon1Bar = 38,

    /// <summary>
    /// 2个条形
    /// </summary>
    xlIcon2Bars = 39,

    /// <summary>
    /// 3个条形
    /// </summary>
    xlIcon3Bars = 40,

    /// <summary>
    /// 4个条形
    /// </summary>
    xlIcon4Bars = 41,

    /// <summary>
    /// 金色星星
    /// </summary>
    xlIconGoldStar = 42,

    /// <summary>
    /// 半颗金色星星
    /// </summary>
    xlIconHalfGoldStar = 43,

    /// <summary>
    /// 银色星星
    /// </summary>
    xlIconSilverStar = 44,

    /// <summary>
    /// 绿色向上三角形
    /// </summary>
    xlIconGreenUpTriangle = 45,

    /// <summary>
    /// 黄色短划线
    /// </summary>
    xlIconYellowDash = 46,

    /// <summary>
    /// 红色向下三角形
    /// </summary>
    xlIconRedDownTriangle = 47,

    /// <summary>
    /// 4个实心方框
    /// </summary>
    xlIcon4FilledBoxes = 48,

    /// <summary>
    /// 3个实心方框
    /// </summary>
    xlIcon3FilledBoxes = 49,

    /// <summary>
    /// 2个实心方框
    /// </summary>
    xlIcon2FilledBoxes = 50,

    /// <summary>
    /// 1个实心方框
    /// </summary>
    xlIcon1FilledBox = 51,

    /// <summary>
    /// 0个实心方框
    /// </summary>
    xlIcon0FilledBoxes = 52
}