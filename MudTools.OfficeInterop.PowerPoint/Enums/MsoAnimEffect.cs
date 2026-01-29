//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定动画效果的枚举。
/// </summary>
public enum MsoAnimEffect
{
    /// <summary>
    /// 自定义动画效果。
    /// </summary>
    msoAnimEffectCustom,

    /// <summary>
    /// 出现动画效果。
    /// </summary>
    msoAnimEffectAppear,

    /// <summary>
    /// 飞入动画效果。
    /// </summary>
    msoAnimEffectFly,

    /// <summary>
    /// 百叶窗动画效果。
    /// </summary>
    msoAnimEffectBlinds,

    /// <summary>
    /// 盒状动画效果。
    /// </summary>
    msoAnimEffectBox,

    /// <summary>
    /// 棋盘动画效果。
    /// </summary>
    msoAnimEffectCheckerboard,

    /// <summary>
    /// 圆形扩展动画效果。
    /// </summary>
    msoAnimEffectCircle,

    /// <summary>
    /// 缓慢移入动画效果。
    /// </summary>
    msoAnimEffectCrawl,

    /// <summary>
    /// 菱形动画效果。
    /// </summary>
    msoAnimEffectDiamond,

    /// <summary>
    /// 溶解动画效果。
    /// </summary>
    msoAnimEffectDissolve,

    /// <summary>
    /// 淡入/淡出动画效果。
    /// </summary>
    msoAnimEffectFade,

    /// <summary>
    /// 闪现一次动画效果。
    /// </summary>
    msoAnimEffectFlashOnce,

    /// <summary>
    /// 窥视动画效果。
    /// </summary>
    msoAnimEffectPeek,

    /// <summary>
    /// 加号动画效果。
    /// </summary>
    msoAnimEffectPlus,

    /// <summary>
    /// 随机条动画效果。
    /// </summary>
    msoAnimEffectRandomBars,

    /// <summary>
    /// 螺旋动画效果。
    /// </summary>
    msoAnimEffectSpiral,

    /// <summary>
    /// 劈裂动画效果。
    /// </summary>
    msoAnimEffectSplit,

    /// <summary>
    /// 伸展动画效果。
    /// </summary>
    msoAnimEffectStretch,

    /// <summary>
    /// 条状动画效果。
    /// </summary>
    msoAnimEffectStrips,

    /// <summary>
    /// 旋转动画效果。
    /// </summary>
    msoAnimEffectSwivel,

    /// <summary>
    /// 楔形动画效果。
    /// </summary>
    msoAnimEffectWedge,

    /// <summary>
    /// 轮子动画效果。
    /// </summary>
    msoAnimEffectWheel,

    /// <summary>
    /// 擦除动画效果。
    /// </summary>
    msoAnimEffectWipe,

    /// <summary>
    /// 缩放动画效果。
    /// </summary>
    msoAnimEffectZoom,

    /// <summary>
    /// 随机效果动画效果。
    /// </summary>
    msoAnimEffectRandomEffects,

    /// <summary>
    /// 回旋动画效果。
    /// </summary>
    msoAnimEffectBoomerang,

    /// <summary>
    /// 弹跳动画效果。
    /// </summary>
    msoAnimEffectBounce,

    /// <summary>
    /// 颜色展示动画效果。
    /// </summary>
    msoAnimEffectColorReveal,

    /// <summary>
    /// 字幕动画效果。
    /// </summary>
    msoAnimEffectCredits,

    /// <summary>
    /// 缓入动画效果。
    /// </summary>
    msoAnimEffectEaseIn,

    /// <summary>
    /// 浮动动画效果。
    /// </summary>
    msoAnimEffectFloat,

    /// <summary>
    /// 放大/缩小并旋转动画效果。
    /// </summary>
    msoAnimEffectGrowAndTurn,

    /// <summary>
    /// 光速动画效果。
    /// </summary>
    msoAnimEffectLightSpeed,

    /// <summary>
    /// 纸风车动画效果。
    /// </summary>
    msoAnimEffectPinwheel,

    /// <summary>
    /// 上升动画效果。
    /// </summary>
    msoAnimEffectRiseUp,

    /// <summary>
    /// 挥鞭式动画效果。
    /// </summary>
    msoAnimEffectSwish,

    /// <summary>
    /// 细线动画效果。
    /// </summary>
    msoAnimEffectThinLine,

    /// <summary>
    /// 展开动画效果。
    /// </summary>
    msoAnimEffectUnfold,

    /// <summary>
    /// 鞭打动画效果。
    /// </summary>
    msoAnimEffectWhip,

    /// <summary>
    /// 上升（直线）动画效果。
    /// </summary>
    msoAnimEffectAscend,

    /// <summary>
    /// 中心旋转动画效果。
    /// </summary>
    msoAnimEffectCenterRevolve,

    /// <summary>
    /// 淡入淡出旋转动画效果。
    /// </summary>
    msoAnimEffectFadedSwivel,

    /// <summary>
    /// 下降动画效果。
    /// </summary>
    msoAnimEffectDescend,

    /// <summary>
    /// 抛射动画效果。
    /// </summary>
    msoAnimEffectSling,

    /// <summary>
    /// 旋转器动画效果。
    /// </summary>
    msoAnimEffectSpinner,

    /// <summary>
    /// 弹性拉伸动画效果。
    /// </summary>
    msoAnimEffectStretchy,

    /// <summary>
    /// 快速滑入动画效果。
    /// </summary>
    msoAnimEffectZip,

    /// <summary>
    /// 向上弧线动画效果。
    /// </summary>
    msoAnimEffectArcUp,

    /// <summary>
    /// 淡入淡出缩放动画效果。
    /// </summary>
    msoAnimEffectFadedZoom,

    /// <summary>
    /// 滑翔动画效果。
    /// </summary>
    msoAnimEffectGlide,

    /// <summary>
    /// 扩展动画效果。
    /// </summary>
    msoAnimEffectExpand,

    /// <summary>
    /// 翻转动画效果。
    /// </summary>
    msoAnimEffectFlip,

    /// <summary>
    /// 闪烁动画效果。
    /// </summary>
    msoAnimEffectShimmer,

    /// <summary>
    /// 折叠动画效果。
    /// </summary>
    msoAnimEffectFold,

    /// <summary>
    /// 更改填充颜色动画效果。
    /// </summary>
    msoAnimEffectChangeFillColor,

    /// <summary>
    /// 更改字体动画效果。
    /// </summary>
    msoAnimEffectChangeFont,

    /// <summary>
    /// 更改字体颜色动画效果。
    /// </summary>
    msoAnimEffectChangeFontColor,

    /// <summary>
    /// 更改字体大小动画效果。
    /// </summary>
    msoAnimEffectChangeFontSize,

    /// <summary>
    /// 更改字体样式动画效果。
    /// </summary>
    msoAnimEffectChangeFontStyle,

    /// <summary>
    /// 放大/缩小动画效果。
    /// </summary>
    msoAnimEffectGrowShrink,

    /// <summary>
    /// 更改线条颜色动画效果。
    /// </summary>
    msoAnimEffectChangeLineColor,

    /// <summary>
    /// 旋转动画效果。
    /// </summary>
    msoAnimEffectSpin,

    /// <summary>
    /// 透明度动画效果。
    /// </summary>
    msoAnimEffectTransparency,

    /// <summary>
    /// 加粗闪现动画效果。
    /// </summary>
    msoAnimEffectBoldFlash,

    /// <summary>
    /// 爆炸动画效果。
    /// </summary>
    msoAnimEffectBlast,

    /// <summary>
    /// 加粗展示动画效果。
    /// </summary>
    msoAnimEffectBoldReveal,

    /// <summary>
    /// 颜色笔刷动画效果。
    /// </summary>
    msoAnimEffectBrushOnColor,

    /// <summary>
    /// 下划线笔刷动画效果。
    /// </summary>
    msoAnimEffectBrushOnUnderline,

    /// <summary>
    /// 颜色混合动画效果。
    /// </summary>
    msoAnimEffectColorBlend,

    /// <summary>
    /// 颜色波动动画效果。
    /// </summary>
    msoAnimEffectColorWave,

    /// <summary>
    /// 互补色动画效果。
    /// </summary>
    msoAnimEffectComplementaryColor,

    /// <summary>
    /// 互补色2动画效果。
    /// </summary>
    msoAnimEffectComplementaryColor2,

    /// <summary>
    /// 对比色动画效果。
    /// </summary>
    msoAnimEffectContrastingColor,

    /// <summary>
    /// 变暗动画效果。
    /// </summary>
    msoAnimEffectDarken,

    /// <summary>
    /// 去色动画效果。
    /// </summary>
    msoAnimEffectDesaturate,

    /// <summary>
    /// 闪光灯动画效果。
    /// </summary>
    msoAnimEffectFlashBulb,

    /// <summary>
    /// 闪烁动画效果。
    /// </summary>
    msoAnimEffectFlicker,

    /// <summary>
    /// 带颜色放大动画效果。
    /// </summary>
    msoAnimEffectGrowWithColor,

    /// <summary>
    /// 变亮动画效果。
    /// </summary>
    msoAnimEffectLighten,

    /// <summary>
    /// 样式强调动画效果。
    /// </summary>
    msoAnimEffectStyleEmphasis,

    /// <summary>
    /// 摇摆动画效果。
    /// </summary>
    msoAnimEffectTeeter,

    /// <summary>
    /// 垂直增长动画效果。
    /// </summary>
    msoAnimEffectVerticalGrow,

    /// <summary>
    /// 波形动画效果。
    /// </summary>
    msoAnimEffectWave,

    /// <summary>
    /// 媒体播放动画效果。
    /// </summary>
    msoAnimEffectMediaPlay,

    /// <summary>
    /// 媒体暂停动画效果。
    /// </summary>
    msoAnimEffectMediaPause,

    /// <summary>
    /// 媒体停止动画效果。
    /// </summary>
    msoAnimEffectMediaStop,

    /// <summary>
    /// 圆形路径动画效果。
    /// </summary>
    msoAnimEffectPathCircle,

    /// <summary>
    /// 直角三角形路径动画效果。
    /// </summary>
    msoAnimEffectPathRightTriangle,

    /// <summary>
    /// 菱形路径动画效果。
    /// </summary>
    msoAnimEffectPathDiamond,

    /// <summary>
    /// 六边形路径动画效果。
    /// </summary>
    msoAnimEffectPathHexagon,

    /// <summary>
    /// 五角星路径动画效果。
    /// </summary>
    msoAnimEffectPath5PointStar,

    /// <summary>
    /// 月牙形路径动画效果。
    /// </summary>
    msoAnimEffectPathCrescentMoon,

    /// <summary>
    /// 正方形路径动画效果。
    /// </summary>
    msoAnimEffectPathSquare,

    /// <summary>
    /// 梯形路径动画效果。
    /// </summary>
    msoAnimEffectPathTrapezoid,

    /// <summary>
    /// 心形路径动画效果。
    /// </summary>
    msoAnimEffectPathHeart,

    /// <summary>
    /// 八边形路径动画效果。
    /// </summary>
    msoAnimEffectPathOctagon,

    /// <summary>
    /// 六角星路径动画效果。
    /// </summary>
    msoAnimEffectPath6PointStar,

    /// <summary>
    /// 橄榄球形路径动画效果。
    /// </summary>
    msoAnimEffectPathFootball,

    /// <summary>
    /// 等边三角形路径动画效果。
    /// </summary>
    msoAnimEffectPathEqualTriangle,

    /// <summary>
    /// 平行四边形路径动画效果。
    /// </summary>
    msoAnimEffectPathParallelogram,

    /// <summary>
    /// 五边形路径动画效果。
    /// </summary>
    msoAnimEffectPathPentagon,

    /// <summary>
    /// 四角星路径动画效果。
    /// </summary>
    msoAnimEffectPath4PointStar,

    /// <summary>
    /// 八角星路径动画效果。
    /// </summary>
    msoAnimEffectPath8PointStar,

    /// <summary>
    /// 泪滴形路径动画效果。
    /// </summary>
    msoAnimEffectPathTeardrop,

    /// <summary>
    /// 尖星形路径动画效果。
    /// </summary>
    msoAnimEffectPathPointyStar,

    /// <summary>
    /// 曲线正方形路径动画效果。
    /// </summary>
    msoAnimEffectPathCurvedSquare,

    /// <summary>
    /// 曲线X形路径动画效果。
    /// </summary>
    msoAnimEffectPathCurvedX,

    /// <summary>
    /// 垂直8字形路径动画效果。
    /// </summary>
    msoAnimEffectPathVerticalFigure8,

    /// <summary>
    /// 曲线星形路径动画效果。
    /// </summary>
    msoAnimEffectPathCurvyStar,

    /// <summary>
    /// 循环路径动画效果。
    /// </summary>
    msoAnimEffectPathLoopdeLoop,

    /// <summary>
    /// 电锯形路径动画效果。
    /// </summary>
    msoAnimEffectPathBuzzsaw,

    /// <summary>
    /// 水平8字形路径动画效果。
    /// </summary>
    msoAnimEffectPathHorizontalFigure8,

    /// <summary>
    /// 花生形路径动画效果。
    /// </summary>
    msoAnimEffectPathPeanut,

    /// <summary>
    /// 四重8字形路径动画效果。
    /// </summary>
    msoAnimEffectPathFigure8Four,

    /// <summary>
    /// 中子形路径动画效果。
    /// </summary>
    msoAnimEffectPathNeutron,

    /// <summary>
    /// 嗖嗖声路径动画效果。
    /// </summary>
    msoAnimEffectPathSwoosh,

    /// <summary>
    /// 豆形路径动画效果。
    /// </summary>
    msoAnimEffectPathBean,

    /// <summary>
    /// 加号形路径动画效果。
    /// </summary>
    msoAnimEffectPathPlus,

    /// <summary>
    /// 倒三角形路径动画效果。
    /// </summary>
    msoAnimEffectPathInvertedTriangle,

    /// <summary>
    /// 倒正方形路径动画效果。
    /// </summary>
    msoAnimEffectPathInvertedSquare,

    /// <summary>
    /// 向左路径动画效果。
    /// </summary>
    msoAnimEffectPathLeft,

    /// <summary>
    /// 向右转弯路径动画效果。
    /// </summary>
    msoAnimEffectPathTurnRight,

    /// <summary>
    /// 向下弧线路径动画效果。
    /// </summary>
    msoAnimEffectPathArcDown,

    /// <summary>
    /// 锯齿形路径动画效果。
    /// </summary>
    msoAnimEffectPathZigzag,

    /// <summary>
    /// 反向S曲线路径动画效果。
    /// </summary>
    msoAnimEffectPathSCurve2,

    /// <summary>
    /// 正弦波路径动画效果。
    /// </summary>
    msoAnimEffectPathSineWave,

    /// <summary>
    /// 向左弹跳路径动画效果。
    /// </summary>
    msoAnimEffectPathBounceLeft,

    /// <summary>
    /// 向下路径动画效果。
    /// </summary>
    msoAnimEffectPathDown,

    /// <summary>
    /// 向上转弯路径动画效果。
    /// </summary>
    msoAnimEffectPathTurnUp,

    /// <summary>
    /// 向上弧线路径动画效果。
    /// </summary>
    msoAnimEffectPathArcUp,

    /// <summary>
    /// 心跳路径动画效果。
    /// </summary>
    msoAnimEffectPathHeartbeat,

    /// <summary>
    /// 右螺旋路径动画效果。
    /// </summary>
    msoAnimEffectPathSpiralRight,

    /// <summary>
    /// 波浪路径动画效果。
    /// </summary>
    msoAnimEffectPathWave,

    /// <summary>
    /// 左曲线路径动画效果。
    /// </summary>
    msoAnimEffectPathCurvyLeft,

    /// <summary>
    /// 对角线向下向右路径动画效果。
    /// </summary>
    msoAnimEffectPathDiagonalDownRight,

    /// <summary>
    /// 向下转弯路径动画效果。
    /// </summary>
    msoAnimEffectPathTurnDown,

    /// <summary>
    /// 向左弧线路径动画效果。
    /// </summary>
    msoAnimEffectPathArcLeft,

    /// <summary>
    /// 漏斗形路径动画效果。
    /// </summary>
    msoAnimEffectPathFunnel,

    /// <summary>
    /// 弹簧形路径动画效果。
    /// </summary>
    msoAnimEffectPathSpring,

    /// <summary>
    /// 向右弹跳路径动画效果。
    /// </summary>
    msoAnimEffectPathBounceRight,

    /// <summary>
    /// 左螺旋路径动画效果。
    /// </summary>
    msoAnimEffectPathSpiralLeft,

    /// <summary>
    /// 对角线向上向右路径动画效果。
    /// </summary>
    msoAnimEffectPathDiagonalUpRight,

    /// <summary>
    /// 向右上方转弯路径动画效果。
    /// </summary>
    msoAnimEffectPathTurnUpRight,

    /// <summary>
    /// 向右弧线路径动画效果。
    /// </summary>
    msoAnimEffectPathArcRight,

    /// <summary>
    /// 正向S曲线路径动画效果。
    /// </summary>
    msoAnimEffectPathSCurve1,

    /// <summary>
    /// 衰减波路径动画效果。
    /// </summary>
    msoAnimEffectPathDecayingWave,

    /// <summary>
    /// 右曲线路径动画效果。
    /// </summary>
    msoAnimEffectPathCurvyRight,

    /// <summary>
    /// 向下阶梯路径动画效果。
    /// </summary>
    msoAnimEffectPathStairsDown,

    /// <summary>
    /// 向上路径动画效果。
    /// </summary>
    msoAnimEffectPathUp,

    /// <summary>
    /// 向右路径动画效果。
    /// </summary>
    msoAnimEffectPathRight,

    /// <summary>
    /// 从书签开始媒体播放动画效果。
    /// </summary>
    msoAnimEffectMediaPlayFromBookmark
}