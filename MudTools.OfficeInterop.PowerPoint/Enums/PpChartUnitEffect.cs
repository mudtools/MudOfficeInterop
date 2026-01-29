//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定图表动画的单元效果。
/// </summary>
public enum PpChartUnitEffect
{
    /// <summary>
    /// 混合图表单元效果。
    /// </summary>
    ppAnimateChartMixed = -2,

    /// <summary>
    /// 按系列设置图表动画。
    /// </summary>
    ppAnimateBySeries = 1,

    /// <summary>
    /// 按类别设置图表动画。
    /// </summary>
    ppAnimateByCategory = 2,

    /// <summary>
    /// 按系列元素设置图表动画。
    /// </summary>
    ppAnimateBySeriesElements = 3,

    /// <summary>
    /// 按类别元素设置图表动画。
    /// </summary>
    ppAnimateByCategoryElements = 4,

    /// <summary>
    /// 一次性显示整个图表动画。
    /// </summary>
    ppAnimateChartAllAtOnce = 5
}