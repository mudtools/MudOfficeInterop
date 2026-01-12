namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定数据透视表切片器的交叉筛选行为类型
/// </summary>
public enum XlSlicerCrossFilterType
{
    /// <summary>
    /// 不应用交叉筛选
    /// </summary>
    xlSlicerNoCrossFilter = 1,
    
    /// <summary>
    /// 交叉筛选时将包含数据的项显示在顶部
    /// </summary>
    xlSlicerCrossFilterShowItemsWithDataAtTop,
    
    /// <summary>
    /// 交叉筛选时显示没有数据的项
    /// </summary>
    xlSlicerCrossFilterShowItemsWithNoData,
    
    /// <summary>
    /// 交叉筛选时隐藏没有数据的按钮
    /// </summary>
    xlSlicerCrossFilterHideButtonsWithNoData
}