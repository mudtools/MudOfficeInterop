namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定数据透视表的数据源类型
/// </summary>
public enum XlPivotTableSourceType
{

    /// <summary>
    /// 数据库查询结果作为数据源
    /// </summary>
    xlDatabase = 1,

    /// <summary>
    /// 外部数据作为数据源
    /// </summary>
    xlExternal = 2,

    /// <summary>
    /// 合并计算数据作为数据源
    /// </summary>
    xlConsolidation = 3,

    /// <summary>
    /// 方案数据作为数据源
    /// </summary>
    xlScenario = 4,

    /// <summary>
    /// 另一个数据透视表作为数据源
    /// </summary>
    xlPivotTable = -4148
}