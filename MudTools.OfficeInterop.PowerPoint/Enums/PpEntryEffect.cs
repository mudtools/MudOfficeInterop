//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;


/// <summary>
/// PowerPoint 进入效果枚举
/// </summary>
public enum PpEntryEffect
{
    /// <summary>
    /// 无效果
    /// </summary>
    ppEffectNone = 0,

    /// <summary>
    /// 切换效果
    /// </summary>
    ppEffectCut = 257,

    /// <summary>
    /// 黑色切出效果
    /// </summary>
    ppEffectCutThroughBlack = 258,

    /// <summary>
    /// 随机效果
    /// </summary>
    ppEffectRandom = 513,

    /// <summary>
    /// 水平百叶窗效果
    /// </summary>
    ppEffectBlindsHorizontal = 769,

    /// <summary>
    /// 垂直百叶窗效果
    /// </summary>
    ppEffectBlindsVertical = 770,

    /// <summary>
    /// 棋盘跨效果
    /// </summary>
    ppEffectCheckerboardAcross = 1025,

    /// <summary>
    /// 棋盘下效果
    /// </summary>
    ppEffectCheckerboardDown = 1026,

    /// <summary>
    /// 从左覆盖效果
    /// </summary>
    ppEffectCoverLeft = 1281,

    /// <summary>
    /// 从上覆盖效果
    /// </summary>
    ppEffectCoverUp = 1282,

    /// <summary>
    /// 从右覆盖效果
    /// </summary>
    ppEffectCoverRight = 1283,

    /// <summary>
    /// 从下覆盖效果
    /// </summary>
    ppEffectCoverDown = 1284,

    /// <summary>
    /// 从左上覆盖效果
    /// </summary>
    ppEffectCoverLeftUp = 1285,

    /// <summary>
    /// 从右上覆盖效果
    /// </summary>
    ppEffectCoverRightUp = 1286,

    /// <summary>
    /// 从左下覆盖效果
    /// </summary>
    ppEffectCoverLeftDown = 1287,

    /// <summary>
    /// 从右下覆盖效果
    /// </summary>
    ppEffectCoverRightDown = 1288,

    /// <summary>
    /// 溶解效果
    /// </summary>
    ppEffectDissolve = 1537,

    /// <summary>
    /// 淡入效果
    /// </summary>
    ppEffectFade = 1793,

    /// <summary>
    /// 快速闪烁一次效果
    /// </summary>
    ppEffectFlashOnceFast = 2049,

    /// <summary>
    /// 中速闪烁一次效果
    /// </summary>
    ppEffectFlashOnceMedium = 2050,

    /// <summary>
    /// 慢速闪烁一次效果
    /// </summary>
    ppEffectFlashOnceSlow = 2051,

    /// <summary>
    /// 从左飞入效果
    /// </summary>
    ppEffectFlyFromLeft = 2305,

    /// <summary>
    /// 从下飞入效果
    /// </summary>
    ppEffectFlyFromBottom = 2306,

    /// <summary>
    /// 从右飞入效果
    /// </summary>
    ppEffectFlyFromRight = 2307,

    /// <summary>
    /// 从上飞入效果
    /// </summary>
    ppEffectFlyFromTop = 2308,

    /// <summary>
    /// 从左上飞入效果
    /// </summary>
    ppEffectFlyFromTopLeft = 2309,

    /// <summary>
    /// 从右上飞入效果
    /// </summary>
    ppEffectFlyFromTopRight = 2310,

    /// <summary>
    /// 从左下飞入效果
    /// </summary>
    ppEffectFlyFromBottomLeft = 2311,

    /// <summary>
    /// 从右下飞入效果
    /// </summary>
    ppEffectFlyFromBottomRight = 2312,

    /// <summary>
    /// 从左切入效果
    /// </summary>
    ppEffectPeekFromLeft = 2561,

    /// <summary>
    /// 从下切入效果
    /// </summary>
    ppEffectPeekFromDown = 2562,

    /// <summary>
    /// 从右切入效果
    /// </summary>
    ppEffectPeekFromRight = 2563,

    /// <summary>
    /// 从上切入效果
    /// </summary>
    ppEffectPeekFromUp = 2564,

    /// <summary>
    /// 加号效果
    /// </summary>
    ppEffectPlus = 2817,

    /// <summary>
    /// 向下推效果
    /// </summary>
    ppEffectPushDown = 3073,

    /// <summary>
    /// 向左推效果
    /// </summary>
    ppEffectPushLeft = 3074,

    /// <summary>
    /// 向右推效果
    /// </summary>
    ppEffectPushRight = 3075,

    /// <summary>
    /// 向上推效果
    /// </summary>
    ppEffectPushUp = 3076,

    /// <summary>
    /// 水平向外分割效果
    /// </summary>
    ppEffectSplitHorizontalOut = 3329,

    /// <summary>
    /// 水平向内分割效果
    /// </summary>
    ppEffectSplitHorizontalIn = 3330,

    /// <summary>
    /// 垂直向外分割效果
    /// </summary>
    ppEffectSplitVerticalOut = 3331,

    /// <summary>
    /// 垂直向内分割效果
    /// </summary>
    ppEffectSplitVerticalIn = 3332,

    /// <summary>
    /// 横向拉伸效果
    /// </summary>
    ppEffectStretchAcross = 3585,

    /// <summary>
    /// 向下拉伸效果
    /// </summary>
    ppEffectStretchDown = 3586,

    /// <summary>
    /// 向左拉伸效果
    /// </summary>
    ppEffectStretchLeft = 3587,

    /// <summary>
    /// 向右拉伸效果
    /// </summary>
    ppEffectStretchRight = 3588,

    /// <summary>
    /// 向上拉伸效果
    /// </summary>
    ppEffectStretchUp = 3589,

    /// <summary>
    /// 从左揭开效果
    /// </summary>
    ppEffectUncoverLeft = 3841,

    /// <summary>
    /// 从上揭开效果
    /// </summary>
    ppEffectUncoverUp = 3842,

    /// <summary>
    /// 从右揭开效果
    /// </summary>
    ppEffectUncoverRight = 3843,

    /// <summary>
    /// 从下揭开效果
    /// </summary>
    ppEffectUncoverDown = 3844,

    /// <summary>
    /// 从左上揭开效果
    /// </summary>
    ppEffectUncoverLeftUp = 3845,

    /// <summary>
    /// 从右上揭开效果
    /// </summary>
    ppEffectUncoverRightUp = 3846,

    /// <summary>
    /// 从左下揭开效果
    /// </summary>
    ppEffectUncoverLeftDown = 3847,

    /// <summary>
    /// 从右下揭开效果
    /// </summary>
    ppEffectUncoverRightDown = 3848,

    /// <summary>
    /// 向下擦除效果
    /// </summary>
    ppEffectWipeDown = 4097,

    /// <summary>
    /// 向左擦除效果
    /// </summary>
    ppEffectWipeLeft = 4098,

    /// <summary>
    /// 向右擦除效果
    /// </summary>
    ppEffectWipeRight = 4099,

    /// <summary>
    /// 向上擦除效果
    /// </summary>
    ppEffectWipeUp = 4100,

    /// <summary>
    /// 放大效果
    /// </summary>
    ppEffectZoomIn = 4353,

    /// <summary>
    /// 轻微放大效果
    /// </summary>
    ppEffectZoomInSlightly = 4354,

    /// <summary>
    /// 缩小效果
    /// </summary>
    ppEffectZoomOut = 4355,

    /// <summary>
    /// 轻微缩小效果
    /// </summary>
    ppEffectZoomOutSlightly = 4356,

    /// <summary>
    /// 中心放大效果
    /// </summary>
    ppEffectZoomCenter = 4357,

    /// <summary>
    /// 底部放大效果
    /// </summary>
    ppEffectZoomBottom = 4358,

    /// <summary>
    /// 顶部放大效果
    /// </summary>
    ppEffectZoomTop = 4359
}