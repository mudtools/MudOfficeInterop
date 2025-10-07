namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel ChartGroups 集合对象的二次封装实现类
/// </summary>
public interface IExcelChartGroups : IEnumerable<IExcelChartGroup>, IDisposable
{

    #region 基础属性

    /// <summary>
    /// 获取集合中图表组的数量
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 通过索引获取集合中的图表组
    /// </summary>
    /// <param name="index">图表组的索引位置</param>
    /// <returns>指定索引位置的图表组</returns>
    IExcelChartGroup? this[int index] { get; }

    /// <summary>
    /// 通过名称获取集合中的图表组
    /// </summary>
    /// <param name="name">图表组的名称</param>
    /// <returns>指定名称的图表组</returns>
    IExcelChartGroup? this[string name] { get; }

    /// <summary>
    /// 获取该对象的父对象
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取应用程序对象
    /// </summary>
    IExcelApplication? Application { get; }
    #endregion
}