//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop;

/// <summary>
/// 指定图表类型
/// </summary>
public enum MsoDiagramType
{
    /// <summary>
    /// 混合图表类型
    /// </summary>
    msoDiagramMixed = -2,

    /// <summary>
    /// 组织结构图
    /// </summary>
    msoDiagramOrgChart = 1,

    /// <summary>
    /// 循环图
    /// </summary>
    msoDiagramCycle = 2,

    /// <summary>
    /// 射线图
    /// </summary>
    msoDiagramRadial = 3,

    /// <summary>
    /// 金字塔图
    /// </summary>
    msoDiagramPyramid = 4,

    /// <summary>
    /// 维恩图
    /// </summary>
    msoDiagramVenn = 5,

    /// <summary>
    /// 目标图
    /// </summary>
    msoDiagramTarget = 6
}
