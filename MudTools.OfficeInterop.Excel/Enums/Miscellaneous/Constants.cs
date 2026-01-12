//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel常量枚举定义集合
/// 包含Excel应用程序中使用的各种常量值定义
/// </summary>
public enum Constants
{
    /// <summary>
    /// 指定所有元素
    /// </summary>
    xlAll = -4104,
    /// <summary>
    /// Excel自动设置
    /// </summary>
    xlAutomatic = -4105,
    /// <summary>
    /// 两边对齐
    /// </summary>
    xlBoth = 1,
    /// <summary>
    /// 居中对齐
    /// </summary>
    xlCenter = -4108,
    /// <summary>
    /// 棋盘格图案样式
    /// </summary>
    xlChecker = 9,
    /// <summary>
    /// 圆形图案样式
    /// </summary>
    xlCircle = 8,
    /// <summary>
    /// 角落位置
    /// </summary>
    xlCorner = 2,
    /// <summary>
    /// 交叉网格线图案样式
    /// </summary>
    xlCrissCross = 16,
    /// <summary>
    /// 十字图案样式
    /// </summary>
    xlCross = 4,
    /// <summary>
    /// 菱形图案样式
    /// </summary>
    xlDiamond = 2,
    /// <summary>
    /// 分散对齐
    /// </summary>
    xlDistributed = -4117,
    /// <summary>
    /// 双会计线样式
    /// </summary>
    xlDoubleAccounting = 5,
    /// <summary>
    /// 固定值
    /// </summary>
    xlFixedValue = 1,
    /// <summary>
    /// 格式设置
    /// </summary>
    xlFormats = -4122,
    /// <summary>
    /// 16%灰色图案
    /// </summary>
    xlGray16 = 17,
    /// <summary>
    /// 8%灰色图案
    /// </summary>
    xlGray8 = 18,
    /// <summary>
    /// 网格线图案
    /// </summary>
    xlGrid = 15,
    /// <summary>
    /// 高位置
    /// </summary>
    xlHigh = -4127,
    /// <summary>
    /// 内部位置
    /// </summary>
    xlInside = 2,
    /// <summary>
    /// 两端对齐
    /// </summary>
    xlJustify = -4130,
    /// <summary>
    /// 向下斜纹图案
    /// </summary>
    xlLightDown = 13,
    /// <summary>
    /// 水平线图案
    /// </summary>
    xlLightHorizontal = 11,
    /// <summary>
    /// 向上斜纹图案
    /// </summary>
    xlLightUp = 14,
    /// <summary>
    /// 垂直线图案
    /// </summary>
    xlLightVertical = 12,
    /// <summary>
    /// 低位置
    /// </summary>
    xlLow = -4134,
    /// <summary>
    /// 手动设置
    /// </summary>
    xlManual = -4135,
    /// <summary>
    /// 负值
    /// </summary>
    xlMinusValues = 3,
    /// <summary>
    /// 模块对象
    /// </summary>
    xlModule = -4141,
    /// <summary>
    /// 紧邻坐标轴
    /// </summary>
    xlNextToAxis = 4,
    /// <summary>
    /// 无设置
    /// </summary>
    xlNone = -4142,
    /// <summary>
    /// 注释
    /// </summary>
    xlNotes = -4144,
    /// <summary>
    /// 关闭状态
    /// </summary>
    xlOff = -4146,
    /// <summary>
    /// 开启状态
    /// </summary>
    xlOn = 1,
    /// <summary>
    /// 百分比
    /// </summary>
    xlPercent = 2,
    /// <summary>
    /// 加号
    /// </summary>
    xlPlus = 9,
    /// <summary>
    /// 正值
    /// </summary>
    xlPlusValues = 2,
    /// <summary>
    /// 75%半灰图案
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
    /// 简单样式
    /// </summary>
    xlSimple = -4154,
    /// <summary>
    /// 单线样式
    /// </summary>
    xlSingle = 2,
    /// <summary>
    /// 单会计线样式
    /// </summary>
    xlSingleAccounting = 4,
    /// <summary>
    /// 实心填充
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
    /// 标准误差
    /// </summary>
    xlStError = 4,
    /// <summary>
    /// 工具栏按钮
    /// </summary>
    xlToolbarButton = 2,
    /// <summary>
    /// 三角形
    /// </summary>
    xlTriangle = 3,
    /// <summary>
    /// 25%灰色
    /// </summary>
    xlGray25 = -4124,
    /// <summary>
    /// 50%灰色
    /// </summary>
    xlGray50 = -4125,
    /// <summary>
    /// 75%灰色
    /// </summary>
    xlGray75 = -4126,
    /// <summary>
    /// 底部对齐
    /// </summary>
    xlBottom = -4107,
    /// <summary>
    /// 左对齐
    /// </summary>
    xlLeft = -4131,
    /// <summary>
    /// 右对齐
    /// </summary>
    xlRight = -4152,
    /// <summary>
    /// 顶部对齐
    /// </summary>
    xlTop = -4160,
    /// <summary>
    /// 3D柱形图
    /// </summary>
    xl3DBar = -4099,
    /// <summary>
    /// 3D曲面图
    /// </summary>
    xl3DSurface = -4103,
    /// <summary>
    /// 柱形图
    /// </summary>
    xlBar = 2,
    /// <summary>
    /// 柱状图
    /// </summary>
    xlColumn = 3,
    /// <summary>
    /// 组合图表
    /// </summary>
    xlCombination = -4111,
    /// <summary>
    /// 自定义格式
    /// </summary>
    xlCustom = -4114,
    /// <summary>
    /// 默认自动格式
    /// </summary>
    xlDefaultAutoFormat = -1,
    /// <summary>
    /// 最大值
    /// </summary>
    xlMaximum = 2,
    /// <summary>
    /// 最小值
    /// </summary>
    xlMinimum = 4,
    /// <summary>
    /// 不透明
    /// </summary>
    xlOpaque = 3,
    /// <summary>
    /// 透明
    /// </summary>
    xlTransparent = 2,
    /// <summary>
    /// 双向文本
    /// </summary>
    xlBidi = -5000,
    /// <summary>
    /// 拉丁文
    /// </summary>
    xlLatin = -5001,
    /// <summary>
    /// 上下文相关
    /// </summary>
    xlContext = -5002,
    /// <summary>
    /// 从左到右
    /// </summary>
    xlLTR = -5003,
    /// <summary>
    /// 从右到左
    /// </summary>
    xlRTL = -5004,
    /// <summary>
    /// 完整脚本
    /// </summary>
    xlFullScript = 1,
    /// <summary>
    /// 部分脚本
    /// </summary>
    xlPartialScript = 2,
    /// <summary>
    /// 混合脚本
    /// </summary>
    xlMixedScript = 3,
    /// <summary>
    /// 混合授权脚本
    /// </summary>
    xlMixedAuthorizedScript = 4,
    /// <summary>
    /// 可视光标
    /// </summary>
    xlVisualCursor = 2,
    /// <summary>
    /// 逻辑光标
    /// </summary>
    xlLogicalCursor = 1,
    /// <summary>
    /// 系统设置
    /// </summary>
    xlSystem = 1,
    /// <summary>
    /// 部分
    /// </summary>
    xlPartial = 3,
    /// <summary>
    /// 印度数字
    /// </summary>
    xlHindiNumerals = 3,
    /// <summary>
    /// 双向日历
    /// </summary>
    xlBidiCalendar = 3,
    /// <summary>
    /// 公历
    /// </summary>
    xlGregorian = 2,
    /// <summary>
    /// 完成状态
    /// </summary>
    xlComplete = 4,
    /// <summary>
    /// 缩放
    /// </summary>
    xlScale = 3,
    /// <summary>
    /// 关闭状态
    /// </summary>
    xlClosed = 3,
    /// <summary>
    /// 颜色1
    /// </summary>
    xlColor1 = 7,
    /// <summary>
    /// 颜色2
    /// </summary>
    xlColor2 = 8,
    /// <summary>
    /// 颜色3
    /// </summary>
    xlColor3 = 9,
    /// <summary>
    /// 常量
    /// </summary>
    xlConstants = 2,
    /// <summary>
    /// 内容
    /// </summary>
    xlContents = 2,
    /// <summary>
    /// 下方
    /// </summary>
    xlBelow = 1,
    /// <summary>
    /// 层叠窗口
    /// </summary>
    xlCascade = 7,
    /// <summary>
    /// 跨选区居中
    /// </summary>
    xlCenterAcrossSelection = 7,
    /// <summary>
    /// 图表4
    /// </summary>
    xlChart4 = 2,
    /// <summary>
    /// 图表系列
    /// </summary>
    xlChartSeries = 17,
    /// <summary>
    /// 短图表
    /// </summary>
    xlChartShort = 6,
    /// <summary>
    /// 图表标题
    /// </summary>
    xlChartTitles = 18,
    /// <summary>
    /// 经典样式1
    /// </summary>
    xlClassic1 = 1,
    /// <summary>
    /// 经典样式2
    /// </summary>
    xlClassic2 = 2,
    /// <summary>
    /// 经典样式3
    /// </summary>
    xlClassic3 = 3,
    /// <summary>
    /// 3D效果1
    /// </summary>
    xl3DEffects1 = 13,
    /// <summary>
    /// 3D效果2
    /// </summary>
    xl3DEffects2 = 14,
    /// <summary>
    /// 上方
    /// </summary>
    xlAbove = 0,
    /// <summary>
    /// 会计格式1
    /// </summary>
    xlAccounting1 = 4,
    /// <summary>
    /// 会计格式2
    /// </summary>
    xlAccounting2 = 5,
    /// <summary>
    /// 会计格式3
    /// </summary>
    xlAccounting3 = 6,
    /// <summary>
    /// 会计格式4
    /// </summary>
    xlAccounting4 = 17,
    /// <summary>
    /// 加法运算
    /// </summary>
    xlAdd = 2,
    /// <summary>
    /// 调试代码窗格
    /// </summary>
    xlDebugCodePane = 13,
    /// <summary>
    /// 桌面
    /// </summary>
    xlDesktop = 9,
    /// <summary>
    /// 直接
    /// </summary>
    xlDirect = 1,
    /// <summary>
    /// 除法运算
    /// </summary>
    xlDivide = 5,
    /// <summary>
    /// 双闭合
    /// </summary>
    xlDoubleClosed = 5,
    /// <summary>
    /// 双开放
    /// </summary>
    xlDoubleOpen = 4,
    /// <summary>
    /// 双引号
    /// </summary>
    xlDoubleQuote = 1,
    /// <summary>
    /// 整个图表
    /// </summary>
    xlEntireChart = 20,
    /// <summary>
    /// Excel菜单
    /// </summary>
    xlExcelMenus = 1,
    /// <summary>
    /// 扩展
    /// </summary>
    xlExtended = 3,
    /// <summary>
    /// 填充
    /// </summary>
    xlFill = 5,
    /// <summary>
    /// 第一个
    /// </summary>
    xlFirst = 0,
    /// <summary>
    /// 浮动
    /// </summary>
    xlFloating = 5,
    /// <summary>
    /// 公式
    /// </summary>
    xlFormula = 5,
    /// <summary>
    /// 常规格式
    /// </summary>
    xlGeneral = 1,
    /// <summary>
    /// 网格线
    /// </summary>
    xlGridline = 22,
    /// <summary>
    /// 图标
    /// </summary>
    xlIcons = 1,
    /// <summary>
    /// 立即窗格
    /// </summary>
    xlImmediatePane = 12,
    /// <summary>
    /// 整数
    /// </summary>
    xlInteger = 2,
    /// <summary>
    /// 最后一个
    /// </summary>
    xlLast = 1,
    /// <summary>
    /// 最后一个单元格
    /// </summary>
    xlLastCell = 11,
    /// <summary>
    /// 列表1
    /// </summary>
    xlList1 = 10,
    /// <summary>
    /// 列表2
    /// </summary>
    xlList2 = 11,
    /// <summary>
    /// 列表3
    /// </summary>
    xlList3 = 12,
    /// <summary>
    /// 本地格式1
    /// </summary>
    xlLocalFormat1 = 15,
    /// <summary>
    /// 本地格式2
    /// </summary>
    xlLocalFormat2 = 16,
    /// <summary>
    /// 长整型
    /// </summary>
    xlLong = 3,
    /// <summary>
    /// Lotus帮助
    /// </summary>
    xlLotusHelp = 2,
    /// <summary>
    /// 宏表单元格
    /// </summary>
    xlMacrosheetCell = 7,
    /// <summary>
    /// 混合状态
    /// </summary>
    xlMixed = 2,
    /// <summary>
    /// 乘法运算
    /// </summary>
    xlMultiply = 4,
    /// <summary>
    /// 窄
    /// </summary>
    xlNarrow = 1,
    /// <summary>
    /// 无文档
    /// </summary>
    xlNoDocuments = 3,
    /// <summary>
    /// 打开
    /// </summary>
    xlOpen = 2,
    /// <summary>
    /// 外部
    /// </summary>
    xlOutside = 3,
    /// <summary>
    /// 引用
    /// </summary>
    xlReference = 4,
    /// <summary>
    /// 半自动
    /// </summary>
    xlSemiautomatic = 2,
    /// <summary>
    /// 短
    /// </summary>
    xlShort = 1,
    /// <summary>
    /// 单引号
    /// </summary>
    xlSingleQuote = 2,
    /// <summary>
    /// 严格
    /// </summary>
    xlStrict = 2,
    /// <summary>
    /// 减法运算
    /// </summary>
    xlSubtract = 3,
    /// <summary>
    /// 文本框
    /// </summary>
    xlTextBox = 16,
    /// <summary>
    /// 平铺
    /// </summary>
    xlTiled = 1,
    /// <summary>
    /// 标题栏
    /// </summary>
    xlTitleBar = 8,
    /// <summary>
    /// 工具栏
    /// </summary>
    xlToolbar = 1,
    /// <summary>
    /// 可见
    /// </summary>
    xlVisible = 12,
    /// <summary>
    /// 监视窗格
    /// </summary>
    xlWatchPane = 11,
    /// <summary>
    /// 宽
    /// </summary>
    xlWide = 3,
    /// <summary>
    /// 工作簿标签
    /// </summary>
    xlWorkbookTab = 6,
    /// <summary>
    /// 工作表4
    /// </summary>
    xlWorksheet4 = 1,
    /// <summary>
    /// 工作表单元格
    /// </summary>
    xlWorksheetCell = 3,
    /// <summary>
    /// 工作表短格式
    /// </summary>
    xlWorksheetShort = 5,
    /// <summary>
    /// 除边框外全部
    /// </summary>
    xlAllExceptBorders = 7,
    /// <summary>
    /// 从左到右
    /// </summary>
    xlLeftToRight = 2,
    /// <summary>
    /// 从上到下
    /// </summary>
    xlTopToBottom = 1,
    /// <summary>
    /// 隐藏
    /// </summary>
    xlVeryHidden = 2,
    /// <summary>
    /// 绘图对象
    /// </summary>
    xlDrawingObject = 14
}