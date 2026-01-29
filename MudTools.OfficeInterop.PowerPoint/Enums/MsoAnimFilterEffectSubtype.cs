//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定滤镜动画效果的子类型。
/// </summary>
public enum MsoAnimFilterEffectSubtype
{
    /// <summary>
    /// 无滤镜动画效果子类型。
    /// </summary>
    msoAnimFilterEffectSubtypeNone,

    /// <summary>
    /// 垂直向内滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeInVertical,

    /// <summary>
    /// 垂直向外滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeOutVertical,

    /// <summary>
    /// 水平向内滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeInHorizontal,

    /// <summary>
    /// 水平向外滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeOutHorizontal,

    /// <summary>
    /// 水平滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeHorizontal,

    /// <summary>
    /// 垂直滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeVertical,

    /// <summary>
    /// 向内滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeIn,

    /// <summary>
    /// 向外滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeOut,

    /// <summary>
    /// 横向穿越滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeAcross,

    /// <summary>
    /// 从左侧滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeFromLeft,

    /// <summary>
    /// 从右侧滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeFromRight,

    /// <summary>
    /// 从顶部滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeFromTop,

    /// <summary>
    /// 从底部滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeFromBottom,

    /// <summary>
    /// 左下角滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeDownLeft,

    /// <summary>
    /// 左上角滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeUpLeft,

    /// <summary>
    /// 右下角滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeDownRight,

    /// <summary>
    /// 右上角滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeUpRight,

    /// <summary>
    /// 1个辐条滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeSpokes1,

    /// <summary>
    /// 2个辐条滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeSpokes2,

    /// <summary>
    /// 3个辐条滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeSpokes3,

    /// <summary>
    /// 4个辐条滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeSpokes4,

    /// <summary>
    /// 8个辐条滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeSpokes8,

    /// <summary>
    /// 向左滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeLeft,

    /// <summary>
    /// 向右滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeRight,

    /// <summary>
    /// 向下滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeDown,

    /// <summary>
    /// 向上滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectSubtypeUp
}