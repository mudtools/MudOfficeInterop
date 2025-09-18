namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定切片器的排序方式
/// </summary>
public enum XlSlicerSort
{
    /// <summary>
    /// 按数据源顺序排序
    /// </summary>
    xlSlicerSortDataSourceOrder = 1,
    
    /// <summary>
    /// 按升序排序
    /// </summary>
    xlSlicerSortAscending,
    
    /// <summary>
    /// 按降序排序
    /// </summary>
    xlSlicerSortDescending
}