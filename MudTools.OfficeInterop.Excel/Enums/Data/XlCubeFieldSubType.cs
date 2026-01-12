//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定 CubeField 的子类型
/// 该枚举用于在处理 OLAP 数据透视表时标识不同类型的 Cube 字段子类型
/// </summary>
public enum XlCubeFieldSubType
{
    /// <summary>
    /// 表示层次结构
    /// </summary>
    xlCubeHierarchy = 1,

    /// <summary>
    /// 表示度量
    /// </summary>
    xlCubeMeasure,

    /// <summary>
    /// 表示集合
    /// </summary>
    xlCubeSet,

    /// <summary>
    /// 表示属性
    /// </summary>
    xlCubeAttribute,

    /// <summary>
    /// 表示计算度量
    /// </summary>
    xlCubeCalculatedMeasure,

    /// <summary>
    /// 表示 KPI 值
    /// </summary>
    xlCubeKPIValue,

    /// <summary>
    /// 表示 KPI 目标
    /// </summary>
    xlCubeKPIGoal,

    /// <summary>
    /// 表示 KPI 状态
    /// </summary>
    xlCubeKPIStatus,

    /// <summary>
    /// 表示 KPI 趋势
    /// </summary>
    xlCubeKPITrend,

    /// <summary>
    /// 表示 KPI 权重
    /// </summary>
    xlCubeKPIWeight,

    /// <summary>
    /// 表示隐式度量
    /// </summary>
    xlCubeImplicitMeasure
}