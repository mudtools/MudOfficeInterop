//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定艺术字文本效果的形状类型
/// </summary>
public enum MsoPresetTextEffectShape
{
    /// <summary>
    /// 混合形状（保留值）
    /// </summary>
    msoTextEffectShapeMixed = -2,
    /// <summary>
    /// 普通文本（无形状效果）
    /// </summary>
    msoTextEffectShapePlainText = 1,
    /// <summary>
    /// 停止标志形状
    /// </summary>
    msoTextEffectShapeStop = 2,
    /// <summary>
    /// 向上的三角形
    /// </summary>
    msoTextEffectShapeTriangleUp = 3,
    /// <summary>
    /// 向下的三角形
    /// </summary>
    msoTextEffectShapeTriangleDown = 4,
    /// <summary>
    /// 向上的V形
    /// </summary>
    msoTextEffectShapeChevronUp = 5,
    /// <summary>
    /// 向下的V形
    /// </summary>
    msoTextEffectShapeChevronDown = 6,
    /// <summary>
    /// 内环形
    /// </summary>
    msoTextEffectShapeRingInside = 7,
    /// <summary>
    /// 外环形
    /// </summary>
    msoTextEffectShapeRingOutside = 8,
    /// <summary>
    /// 上拱曲线
    /// </summary>
    msoTextEffectShapeArchUpCurve = 9,
    /// <summary>
    /// 下拱曲线
    /// </summary>
    msoTextEffectShapeArchDownCurve = 10,
    /// <summary>
    /// 圆形曲线
    /// </summary>
    msoTextEffectShapeCircleCurve = 11,
    /// <summary>
    /// 按钮曲线
    /// </summary>
    msoTextEffectShapeButtonCurve = 12,
    /// <summary>
    /// 上拱倾泻形
    /// </summary>
    msoTextEffectShapeArchUpPour = 13,
    /// <summary>
    /// 下拱倾泻形
    /// </summary>
    msoTextEffectShapeArchDownPour = 14,
    /// <summary>
    /// 圆形倾泻形
    /// </summary>
    msoTextEffectShapeCirclePour = 15,
    /// <summary>
    /// 按钮倾泻形
    /// </summary>
    msoTextEffectShapeButtonPour = 16,
    /// <summary>
    /// 向上曲线
    /// </summary>
    msoTextEffectShapeCurveUp = 17,
    /// <summary>
    /// 向下曲线
    /// </summary>
    msoTextEffectShapeCurveDown = 18,
    /// <summary>
    /// 凸面朝上的罐形
    /// </summary>
    msoTextEffectShapeCanUp = 19,
    /// <summary>
    /// 凸面朝下的罐形
    /// </summary>
    msoTextEffectShapeCanDown = 20,
    /// <summary>
    /// 波浪形1
    /// </summary>
    msoTextEffectShapeWave1 = 21,
    /// <summary>
    /// 波浪形2
    /// </summary>
    msoTextEffectShapeWave2 = 22,
    /// <summary>
    /// 双波浪形1
    /// </summary>
    msoTextEffectShapeDoubleWave1 = 23,
    /// <summary>
    /// 双波浪形2
    /// </summary>
    msoTextEffectShapeDoubleWave2 = 24,
    /// <summary>
    /// 膨胀形
    /// </summary>
    msoTextEffectShapeInflate = 25,
    /// <summary>
    /// 压缩形
    /// </summary>
    msoTextEffectShapeDeflate = 26,
    /// <summary>
    /// 底部膨胀形
    /// </summary>
    msoTextEffectShapeInflateBottom = 27,
    /// <summary>
    /// 底部压缩形
    /// </summary>
    msoTextEffectShapeDeflateBottom = 28,
    /// <summary>
    /// 顶部膨胀形
    /// </summary>
    msoTextEffectShapeInflateTop = 29,
    /// <summary>
    /// 顶部压缩形
    /// </summary>
    msoTextEffectShapeDeflateTop = 30,
    /// <summary>
    /// 压缩-膨胀形
    /// </summary>
    msoTextEffectShapeDeflateInflate = 31,
    /// <summary>
    /// 压缩-膨胀-压缩形
    /// </summary>
    msoTextEffectShapeDeflateInflateDeflate = 32,
    /// <summary>
    /// 向右淡出
    /// </summary>
    msoTextEffectShapeFadeRight = 33,
    /// <summary>
    /// 向左淡出
    /// </summary>
    msoTextEffectShapeFadeLeft = 34,
    /// <summary>
    /// 向上淡出
    /// </summary>
    msoTextEffectShapeFadeUp = 35,
    /// <summary>
    /// 向下淡出
    /// </summary>
    msoTextEffectShapeFadeDown = 36,
    /// <summary>
    /// 向上倾斜
    /// </summary>
    msoTextEffectShapeSlantUp = 37,
    /// <summary>
    /// 向下倾斜
    /// </summary>
    msoTextEffectShapeSlantDown = 38,
    /// <summary>
    /// 向上级联
    /// </summary>
    msoTextEffectShapeCascadeUp = 39,
    /// <summary>
    /// 向下级联
    /// </summary>
    msoTextEffectShapeCascadeDown = 40
}