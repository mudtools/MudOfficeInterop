//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定自动形状类型的枚举，用于定义各种几何形状
/// </summary>
public enum MsoAutoShapeType
{
    /// <summary>
    /// 混合形状类型
    /// </summary>
    msoShapeMixed = -2,

    /// <summary>
    /// 矩形
    /// </summary>
    msoShapeRectangle = 1,

    /// <summary>
    /// 平行四边形
    /// </summary>
    msoShapeParallelogram = 2,

    /// <summary>
    /// 梯形
    /// </summary>
    msoShapeTrapezoid = 3,

    /// <summary>
    /// 菱形
    /// </summary>
    msoShapeDiamond = 4,

    /// <summary>
    /// 圆角矩形
    /// </summary>
    msoShapeRoundedRectangle = 5,

    /// <summary>
    /// 八边形
    /// </summary>
    msoShapeOctagon = 6,

    /// <summary>
    /// 等腰三角形
    /// </summary>
    msoShapeIsoscelesTriangle = 7,

    /// <summary>
    /// 直角三角形
    /// </summary>
    msoShapeRightTriangle = 8,

    /// <summary>
    /// 椭圆
    /// </summary>
    msoShapeOval = 9,

    /// <summary>
    /// 六边形
    /// </summary>
    msoShapeHexagon = 10,

    /// <summary>
    /// 十字形
    /// </summary>
    msoShapeCross = 11,

    /// <summary>
    /// 正五边形
    /// </summary>
    msoShapeRegularPentagon = 12,

    /// <summary>
    /// 圆柱形
    /// </summary>
    msoShapeCan = 13,

    /// <summary>
    /// 立方体
    /// </summary>
    msoShapeCube = 14,

    /// <summary>
    /// 倒角
    /// </summary>
    msoShapeBevel = 15,

    /// <summary>
    /// 折角矩形
    /// </summary>
    msoShapeFoldedCorner = 16,

    /// <summary>
    /// 笑脸
    /// </summary>
    msoShapeSmileyFace = 17,

    /// <summary>
    /// 圆环
    /// </summary>
    msoShapeDonut = 18,

    /// <summary>
    /// 禁止符
    /// </summary>
    msoShapeNoSymbol = 19,

    /// <summary>
    /// 弧形
    /// </summary>
    msoShapeBlockArc = 20,

    /// <summary>
    /// 心形
    /// </summary>
    msoShapeHeart = 21,

    /// <summary>
    /// 闪电
    /// </summary>
    msoShapeLightningBolt = 22,

    /// <summary>
    /// 太阳形
    /// </summary>
    msoShapeSun = 23,

    /// <summary>
    /// 月牙形
    /// </summary>
    msoShapeMoon = 24,

    /// <summary>
    /// 弓形
    /// </summary>
    msoShapeArc = 25,

    /// <summary>
    /// 双括号
    /// </summary>
    msoShapeDoubleBracket = 26,

    /// <summary>
    /// 双大括号
    /// </summary>
    msoShapeDoubleBrace = 27,

    /// <summary>
    /// 铭牌形
    /// </summary>
    msoShapePlaque = 28,

    /// <summary>
    /// 人字形
    /// </summary>
    msoShapeChevron = 29,

    /// <summary>
    /// 圆形箭头
    /// </summary>
    msoShapeCircularArrow = 30,

    /// <summary>
    /// 带缺口的圆形箭头
    /// </summary>
    msoShapeNotchedCircularArrow = 31,

    /// <summary>
    /// U形箭头
    /// </summary>
    msoShapeUturnArrow = 32,

    /// <summary>
    /// 向右弯曲箭头
    /// </summary>
    msoShapeCurvedRightArrow = 33,

    /// <summary>
    /// 向左弯曲箭头
    /// </summary>
    msoShapeCurvedLeftArrow = 34,

    /// <summary>
    /// 向上弯曲箭头
    /// </summary>
    msoShapeCurvedUpArrow = 35,

    /// <summary>
    /// 向下弯曲箭头
    /// </summary>
    msoShapeCurvedDownArrow = 36,

    /// <summary>
    /// 条纹向右箭头
    /// </summary>
    msoShapeStripedRightArrow = 37,

    /// <summary>
    /// 带缺口的向右箭头
    /// </summary>
    msoShapeNotchedRightArrow = 38,

    /// <summary>
    /// 五边形
    /// </summary>
    msoShapePentagon = 39,

    /// <summary>
    /// V形箭头
    /// </summary>
    msoShapeChevronArrow = 40,

    /// <summary>
    /// 六齿齿轮
    /// </summary>
    msoShapeGear6 = 41,

    /// <summary>
    /// 九齿齿轮
    /// </summary>
    msoShapeGear9 = 42,

    /// <summary>
    /// 漏斗形
    /// </summary>
    msoShapeFunnel = 43,

    /// <summary>
    /// 扇形
    /// </summary>
    msoShapePieWedge = 44,

    /// <summary>
    /// 半框架
    /// </summary>
    msoShapeHalfFrame = 45,

    /// <summary>
    /// 对角条纹
    /// </summary>
    msoShapeDiagonalStripe = 46,

    /// <summary>
    /// 饼形
    /// </summary>
    msoShapePie = 47,

    /// <summary>
    /// 不等边梯形
    /// </summary>
    msoShapeNonIsoscelesTrapezoid = 48,

    /// <summary>
    /// 十边形
    /// </summary>
    msoShapeDecagon = 49,

    /// <summary>
    /// 七边形
    /// </summary>
    msoShapeHeptagon = 50,

    /// <summary>
    /// 十一边形
    /// </summary>
    msoShapeHendecagon = 51,

    /// <summary>
    /// 十二边形
    /// </summary>
    msoShapeDodecagon = 52,

    /// <summary>
    /// 六角星
    /// </summary>
    msoShapeStar6Point = 53,

    /// <summary>
    /// 七角星
    /// </summary>
    msoShapeStar7Point = 54,

    /// <summary>
    /// 十角星
    /// </summary>
    msoShapeStar10Point = 55,

    /// <summary>
    /// 十二角星
    /// </summary>
    msoShapeStar12Point = 56,

    /// <summary>
    /// 十六角星
    /// </summary>
    msoShapeStar16Point = 57,

    /// <summary>
    /// 二十四角星
    /// </summary>
    msoShapeStar24Point = 58,

    /// <summary>
    /// 三十二角星
    /// </summary>
    msoShapeStar32Point = 59,

    /// <summary>
    /// 圆角矩形
    /// </summary>
    msoShapeRoundRectangle = 60,

    /// <summary>
    /// 圆角圆角矩形
    /// </summary>
    msoShapeSnipRoundRectangle = 61,

    /// <summary>
    /// 切角矩形（单角）
    /// </summary>
    msoShapeSnipSingleCornerRectangle = 62,

    /// <summary>
    /// 切角矩形（对角）
    /// </summary>
    msoShapeSnipDiagonalCornerRectangle = 63,

    /// <summary>
    /// 铭牌标签
    /// </summary>
    msoShapePlaqueTabs = 64,

    /// <summary>
    /// 方形标签
    /// </summary>
    msoShapeSquareTabs = 65,

    /// <summary>
    /// 七齿齿轮
    /// </summary>
    msoShapeGear7 = 66,

    /// <summary>
    /// 八齿齿轮
    /// </summary>
    msoShapeGear8 = 67,

    /// <summary>
    /// 弦形
    /// </summary>
    msoShapeChord = 68,

    /// <summary>
    /// 角形
    /// </summary>
    msoShapeCorner = 69,

    /// <summary>
    /// 加号
    /// </summary>
    msoShapeMathPlus = 70,

    /// <summary>
    /// 减号
    /// </summary>
    msoShapeMathMinus = 71,

    /// <summary>
    /// 乘号
    /// </summary>
    msoShapeMathMultiply = 72,

    /// <summary>
    /// 除号
    /// </summary>
    msoShapeMathDivide = 73,

    /// <summary>
    /// 等号
    /// </summary>
    msoShapeMathEqual = 74,

    /// <summary>
    /// 不等号
    /// </summary>
    msoShapeMathNotEqual = 75,

    /// <summary>
    /// 角标签
    /// </summary>
    msoShapeCornerTabs = 76,

    /// <summary>
    /// 方形箭头
    /// </summary>
    msoShapeSquareArrow = 77,

    /// <summary>
    /// 本垒板形
    /// </summary>
    msoShapeHomePlate = 78,

    /// <summary>
    /// 宽V形
    /// </summary>
    msoShapeChevronWide = 79,

    /// <summary>
    /// 倒V形
    /// </summary>
    msoShapeChevronInverse = 80,

    /// <summary>
    /// 窄条纹向右箭头
    /// </summary>
    msoShapeStripedRightArrowSlim = 81,

    /// <summary>
    /// 向右标注箭头
    /// </summary>
    msoShapeRightArrowCallout = 82,

    /// <summary>
    /// 向下标注箭头
    /// </summary>
    msoShapeDownArrowCallout = 83,

    /// <summary>
    /// 向左标注箭头
    /// </summary>
    msoShapeLeftArrowCallout = 84,

    /// <summary>
    /// 向上标注箭头
    /// </summary>
    msoShapeUpArrowCallout = 85,

    /// <summary>
    /// 向右弯曲标注箭头
    /// </summary>
    msoShapeCurvedRightArrowCallout = 86,

    /// <summary>
    /// 向下弯曲标注箭头
    /// </summary>
    msoShapeCurvedDownArrowCallout = 87,

    /// <summary>
    /// 向左弯曲标注箭头
    /// </summary>
    msoShapeCurvedLeftArrowCallout = 88,

    /// <summary>
    /// 向上弯曲标注箭头
    /// </summary>
    msoShapeCurvedUpArrowCallout = 89,

    /// <summary>
    /// 标注气泡1
    /// </summary>
    msoShapeCallout1 = 90,

    /// <summary>
    /// 标注气泡2
    /// </summary>
    msoShapeCallout2 = 91,

    /// <summary>
    /// 标注气泡3
    /// </summary>
    msoShapeCallout3 = 92,

    /// <summary>
    /// 标注气泡4
    /// </summary>
    msoShapeCallout4 = 93,

    /// <summary>
    /// 标注气泡5
    /// </summary>
    msoShapeCallout5 = 94,

    /// <summary>
    /// 标注气泡6
    /// </summary>
    msoShapeCallout6 = 95,

    /// <summary>
    /// 标注气泡7
    /// </summary>
    msoShapeCallout7 = 96,

    /// <summary>
    /// 标注气泡8
    /// </summary>
    msoShapeCallout8 = 97,

    /// <summary>
    /// 标注气泡9
    /// </summary>
    msoShapeCallout9 = 98,

    /// <summary>
    /// 强调标注气泡1
    /// </summary>
    msoShapeAccentCallout1 = 99,

    /// <summary>
    /// 强调标注气泡2
    /// </summary>
    msoShapeAccentCallout2 = 100,

    /// <summary>
    /// 强调标注气泡3
    /// </summary>
    msoShapeAccentCallout3 = 101,

    /// <summary>
    /// 强调标注气泡4
    /// </summary>
    msoShapeAccentCallout4 = 102,

    /// <summary>
    /// 强调标注气泡5
    /// </summary>
    msoShapeAccentCallout5 = 103,

    /// <summary>
    /// 强调标注气泡6
    /// </summary>
    msoShapeAccentCallout6 = 104,

    /// <summary>
    /// 强调标注气泡7
    /// </summary>
    msoShapeAccentCallout7 = 105,

    /// <summary>
    /// 强调标注气泡8
    /// </summary>
    msoShapeAccentCallout8 = 106,

    /// <summary>
    /// 强调标注气泡9
    /// </summary>
    msoShapeAccentCallout9 = 107,

    /// <summary>
    /// 边框标注气泡1
    /// </summary>
    msoShapeBorderCallout1 = 108,

    /// <summary>
    /// 边框标注气泡2
    /// </summary>
    msoShapeBorderCallout2 = 109,

    /// <summary>
    /// 边框标注气泡3
    /// </summary>
    msoShapeBorderCallout3 = 110,

    /// <summary>
    /// 边框标注气泡4
    /// </summary>
    msoShapeBorderCallout4 = 111,

    /// <summary>
    /// 边框标注气泡5
    /// </summary>
    msoShapeBorderCallout5 = 112,

    /// <summary>
    /// 边框标注气泡6
    /// </summary>
    msoShapeBorderCallout6 = 113,

    /// <summary>
    /// 边框标注气泡7
    /// </summary>
    msoShapeBorderCallout7 = 114,

    /// <summary>
    /// 边框标注气泡8
    /// </summary>
    msoShapeBorderCallout8 = 115,

    /// <summary>
    /// 边框标注气泡9
    /// </summary>
    msoShapeBorderCallout9 = 116,

    /// <summary>
    /// 丝带形
    /// </summary>
    msoShapeRibbon = 117,

    /// <summary>
    /// 丝带形2
    /// </summary>
    msoShapeRibbon2 = 118,

    /// <summary>
    /// 飞旋箭头
    /// </summary>
    msoShapeSwooshArrow = 119,

    /// <summary>
    /// 泪珠形
    /// </summary>
    msoShapeTeardrop = 120,

    /// <summary>
    /// 主机控件
    /// </summary>
    msoShapeHostControl = 121,

    /// <summary>
    /// 垂直卷轴
    /// </summary>
    msoShapeVerticalScroll = 122,

    /// <summary>
    /// 水平卷轴
    /// </summary>
    msoShapeHorizontalScroll = 123,

    /// <summary>
    /// 波浪形
    /// </summary>
    msoShapeWave = 124,

    /// <summary>
    /// 双波浪形
    /// </summary>
    msoShapeDoubleWave = 125,

    /// <summary>
    /// 矩形标注
    /// </summary>
    msoShapeRectangularCallout = 126,

    /// <summary>
    /// 圆角矩形标注
    /// </summary>
    msoShapeRoundedRectangularCallout = 127,

    /// <summary>
    /// 椭圆标注
    /// </summary>
    msoShapeOvalCallout = 128,

    /// <summary>
    /// 云朵标注
    /// </summary>
    msoShapeCloudCallout = 129,

    /// <summary>
    /// 线形标注1
    /// </summary>
    msoShapeLineCallout1 = 130,

    /// <summary>
    /// 线形标注2
    /// </summary>
    msoShapeLineCallout2 = 131,

    /// <summary>
    /// 线形标注3
    /// </summary>
    msoShapeLineCallout3 = 132,

    /// <summary>
    /// 线形标注4
    /// </summary>
    msoShapeLineCallout4 = 133,

    /// <summary>
    /// 线形标注1（强调）
    /// </summary>
    msoShapeLineCallout1Accent = 134,

    /// <summary>
    /// 线形标注2（强调）
    /// </summary>
    msoShapeLineCallout2Accent = 135,

    /// <summary>
    /// 线形标注3（强调）
    /// </summary>
    msoShapeLineCallout3Accent = 136,

    /// <summary>
    /// 线形标注4（强调）
    /// </summary>
    msoShapeLineCallout4Accent = 137,

    /// <summary>
    /// 线形标注1（无边框）
    /// </summary>
    msoShapeLineCallout1NoBorder = 138,

    /// <summary>
    /// 线形标注2（无边框）
    /// </summary>
    msoShapeLineCallout2NoBorder = 139,

    /// <summary>
    /// 线形标注3（无边框）
    /// </summary>
    msoShapeLineCallout3NoBorder = 140,

    /// <summary>
    /// 线形标注4（无边框）
    /// </summary>
    msoShapeLineCallout4NoBorder = 141,

    /// <summary>
    /// 线形标注1（边框和强调）
    /// </summary>
    msoShapeLineCallout1BorderandAccent = 142,

    /// <summary>
    /// 线形标注2（边框和强调）
    /// </summary>
    msoShapeLineCallout2BorderandAccent = 143,

    /// <summary>
    /// 线形标注3（边框和强调）
    /// </summary>
    msoShapeLineCallout3BorderandAccent = 144,

    /// <summary>
    /// 线形标注4（边框和强调）
    /// </summary>
    msoShapeLineCallout4BorderandAccent = 145
}