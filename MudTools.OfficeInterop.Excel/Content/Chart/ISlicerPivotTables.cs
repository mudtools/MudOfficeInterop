
namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// Excel SlicerPivotTables 集合对象的二次封装实现类
/// </summary>
[ComCollectionWrap(ComNamespace = "MsExcel")]
public interface ISlicerPivotTables : IOfficeObject<ISlicerPivotTables>, IEnumerable<IExcelPivotTable>, IDisposable
{
    /// <summary>
    /// 获取数据透视表集合所在的父对象（通常是 Worksheet）
    /// 对应 PivotTables.Parent 属性
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取数据透视表集合所在的Application对象
    /// 对应 PivotTables.Application 属性
    /// </summary>
    [ComPropertyWrap(NeedDispose = false)]
    IExcelApplication? Application { get; }

    #region 基础属性
    /// <summary>
    /// 获取数据透视表集合中的透视表数量
    /// 对应 PivotTables.Count 属性
    /// </summary>
    int Count { get; }

    /// <summary>
    /// 获取指定索引的数据透视表对象
    /// 索引从1开始
    /// </summary>
    /// <param name="index">透视表索引（从1开始）</param>
    /// <returns>数据透视表对象</returns>
    IExcelPivotTable? this[int index] { get; }

    /// <summary>
    /// 获取指定名称的数据透视表对象
    /// </summary>
    /// <param name="name">透视表名称</param>
    /// <returns>数据透视表对象</returns>
    IExcelPivotTable? this[string name] { get; }

    #endregion
}