//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.PowerPoint;

/// <summary>
/// 指定按级别进行动画的方式。
/// </summary>
public enum MsoAnimateByLevel
{
    /// <summary>
    /// 混合级别动画。
    /// </summary>
    msoAnimateLevelMixed = -1,

    /// <summary>
    /// 无级别动画。
    /// </summary>
    msoAnimateLevelNone,

    /// <summary>
    /// 按所有级别设置文本动画。
    /// </summary>
    msoAnimateTextByAllLevels,

    /// <summary>
    /// 按第一级别设置文本动画。
    /// </summary>
    msoAnimateTextByFirstLevel,

    /// <summary>
    /// 按第二级别设置文本动画。
    /// </summary>
    msoAnimateTextBySecondLevel,

    /// <summary>
    /// 按第三级别设置文本动画。
    /// </summary>
    msoAnimateTextByThirdLevel,

    /// <summary>
    /// 按第四级别设置文本动画。
    /// </summary>
    msoAnimateTextByFourthLevel,

    /// <summary>
    /// 按第五级别设置文本动画。
    /// </summary>
    msoAnimateTextByFifthLevel,

    /// <summary>
    /// 同时设置整个图表动画。
    /// </summary>
    msoAnimateChartAllAtOnce,

    /// <summary>
    /// 按类别设置图表动画。
    /// </summary>
    msoAnimateChartByCategory,

    /// <summary>
    /// 按类别元素设置图表动画。
    /// </summary>
    msoAnimateChartByCategoryElements,

    /// <summary>
    /// 按系列设置图表动画。
    /// </summary>
    msoAnimateChartBySeries,

    /// <summary>
    /// 按系列元素设置图表动画。
    /// </summary>
    msoAnimateChartBySeriesElements,

    /// <summary>
    /// 同时设置整个图表动画。
    /// </summary>
    msoAnimateDiagramAllAtOnce,

    /// <summary>
    /// 按节点深度设置图表动画。
    /// </summary>
    msoAnimateDiagramDepthByNode,

    /// <summary>
    /// 按分支深度设置图表动画。
    /// </summary>
    msoAnimateDiagramDepthByBranch,

    /// <summary>
    /// 按节点广度设置图表动画。
    /// </summary>
    msoAnimateDiagramBreadthByNode,

    /// <summary>
    /// 按级别广度设置图表动画。
    /// </summary>
    msoAnimateDiagramBreadthByLevel,

    /// <summary>
    /// 顺时针设置图表动画。
    /// </summary>
    msoAnimateDiagramClockwise,

    /// <summary>
    /// 顺时针向内设置图表动画。
    /// </summary>
    msoAnimateDiagramClockwiseIn,

    /// <summary>
    /// 顺时针向外设置图表动画。
    /// </summary>
    msoAnimateDiagramClockwiseOut,

    /// <summary>
    /// 逆时针设置图表动画。
    /// </summary>
    msoAnimateDiagramCounterClockwise,

    /// <summary>
    /// 逆时针向内设置图表动画。
    /// </summary>
    msoAnimateDiagramCounterClockwiseIn,

    /// <summary>
    /// 逆时针向外设置图表动画。
    /// </summary>
    msoAnimateDiagramCounterClockwiseOut,

    /// <summary>
    /// 按环形向内设置图表动画。
    /// </summary>
    msoAnimateDiagramInByRing,

    /// <summary>
    /// 按环形向外设置图表动画。
    /// </summary>
    msoAnimateDiagramOutByRing,

    /// <summary>
    /// 向上设置图表动画。
    /// </summary>
    msoAnimateDiagramUp,

    /// <summary>
    /// 向下设置图表动画。
    /// </summary>
    msoAnimateDiagramDown
}