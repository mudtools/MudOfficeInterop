//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定组织结构图的布局类型
/// </summary>
public enum MsoOrgChartLayoutType
{
    /// <summary>
    /// 混合布局类型
    /// </summary>
    msoOrgChartLayoutMixed = -2,

    /// <summary>
    /// 标准布局类型
    /// </summary>
    msoOrgChartLayoutStandard = 1,

    /// <summary>
    /// 两侧悬挂布局类型
    /// </summary>
    msoOrgChartLayoutBothHanging = 2,

    /// <summary>
    /// 左侧悬挂布局类型
    /// </summary>
    msoOrgChartLayoutLeftHanging = 3,

    /// <summary>
    /// 右侧悬挂布局类型
    /// </summary>
    msoOrgChartLayoutRightHanging = 4,
    
    /// <summary>
    /// 默认布局类型
    /// </summary>
    msoOrgChartLayoutDefault = 5
}
