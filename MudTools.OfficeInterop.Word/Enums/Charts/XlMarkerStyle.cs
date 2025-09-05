//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 指定图表数据点或数据系列的标记样式
/// </summary>
public enum XlMarkerStyle
{
    /// <summary>
    /// 自动标记样式
    /// </summary>
    xlMarkerStyleAutomatic = -4105,
    /// <summary>
    /// 圆形标记
    /// </summary>
    xlMarkerStyleCircle = 8,
    /// <summary>
    /// 破折号标记
    /// </summary>
    xlMarkerStyleDash = -4115,
    /// <summary>
    /// 菱形标记
    /// </summary>
    xlMarkerStyleDiamond = 2,
    /// <summary>
    /// 点状标记
    /// </summary>
    xlMarkerStyleDot = -4118,
    /// <summary>
    /// 无标记
    /// </summary>
    xlMarkerStyleNone = -4142,
    /// <summary>
    /// 图片标记
    /// </summary>
    xlMarkerStylePicture = -4147,
    /// <summary>
    /// 加号标记
    /// </summary>
    xlMarkerStylePlus = 9,
    /// <summary>
    /// 正方形标记
    /// </summary>
    xlMarkerStyleSquare = 1,
    /// <summary>
    /// 星形标记
    /// </summary>
    xlMarkerStyleStar = 5,
    /// <summary>
    /// 三角形标记
    /// </summary>
    xlMarkerStyleTriangle = 3,
    /// <summary>
    /// X形标记
    /// </summary>
    xlMarkerStyleX = -4168
}