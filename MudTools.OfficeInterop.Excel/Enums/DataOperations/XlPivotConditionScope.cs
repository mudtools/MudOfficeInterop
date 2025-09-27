namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定数据透视表条件格式的作用范围
/// </summary>
public enum XlPivotConditionScope
{
    /// <summary>
    /// 选择区域范围
    /// </summary>
    xlSelectionScope,
    /// <summary>
    /// 字段范围
    /// </summary>
    xlFieldsScope,
    /// <summary>
    /// 数据字段范围
    /// </summary>
    xlDataFieldScope
}