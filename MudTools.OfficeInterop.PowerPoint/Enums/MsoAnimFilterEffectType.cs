//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定滤镜动画效果的类型。
/// </summary>
public enum MsoAnimFilterEffectType
{
    /// <summary>
    /// 无滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeNone,

    /// <summary>
    /// 谷仓门滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeBarn,

    /// <summary>
    /// 百叶窗滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeBlinds,

    /// <summary>
    /// 盒状滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeBox,

    /// <summary>
    /// 棋盘滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeCheckerboard,

    /// <summary>
    /// 圆形滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeCircle,

    /// <summary>
    /// 菱形滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeDiamond,

    /// <summary>
    /// 溶解滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeDissolve,

    /// <summary>
    /// 淡入淡出滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeFade,

    /// <summary>
    /// 图像滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeImage,

    /// <summary>
    /// 像素化滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypePixelate,

    /// <summary>
    /// 加号滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypePlus,

    /// <summary>
    /// 随机条滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeRandomBar,

    /// <summary>
    /// 滑动滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeSlide,

    /// <summary>
    /// 伸展滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeStretch,

    /// <summary>
    /// 条状滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeStrips,

    /// <summary>
    /// 楔形滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeWedge,

    /// <summary>
    /// 轮子滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeWheel,

    /// <summary>
    /// 擦除滤镜动画效果。
    /// </summary>
    msoAnimFilterEffectTypeWipe
}