namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 指定数据透视表的版本列表
/// </summary>
public enum XlPivotTableVersionList
{
    /// <summary>
    /// Excel 2000版本的数据透视表 (Version 9.0)
    /// </summary>
    xlPivotTableVersion2000 = 0,

    /// <summary>
    /// Excel 2002版本的数据透视表 (Version 10.0)
    /// </summary>
    xlPivotTableVersion10 = 1,
    //

    /// <summary>
    /// Excel 2003版本的数据透视表 (Version 11.0)
    /// </summary>
    xlPivotTableVersion11 = 2,

    /// <summary>
    /// Excel 2007版本的数据透视表 (Version 12.0)
    /// </summary>
    xlPivotTableVersion12 = 3,

    /// <summary>
    /// Excel 2010版本的数据透视表 (Version 14.0)
    /// </summary>
    xlPivotTableVersion14 = 4,

    /// <summary>
    /// Excel 2013版本的数据透视表 (Version 15.0)
    /// </summary>
    xlPivotTableVersion15 = 5,

    /// <summary>
    /// 当前版本的数据透视表
    /// </summary>
    xlPivotTableVersionCurrent = -1
}