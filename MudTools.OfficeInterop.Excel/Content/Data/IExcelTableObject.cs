namespace MudTools.OfficeInterop.Excel;

/// <summary>
/// 表示 Excel 中一个结构化数据表（Table）的封装接口。
/// 对应 COM 对象：Microsoft.Office.Interop.Excel.ListObject
/// 本接口语义命名为 TableObject，实际封装 ListObject，以符合现代 Excel 表格功能。
/// 支持表头、数据体、汇总行、排序、筛选、列管理等。
/// </summary>
public interface IExcelTableObject : IDisposable
{
    // ========== 基础属性 ==========

    /// <summary>
    /// 获取此对象的父对象（通常是 Worksheet）。
    /// </summary>
    object? Parent { get; }

    /// <summary>
    /// 获取此对象所属的 Excel 应用程序对象。
    /// </summary>
    IExcelApplication? Application { get; }


    // ========== 表格行为控制 ==========

    /// <summary>
    /// 是否显示行号。
    /// </summary>
    bool RowNumbers { get; set; }

    /// <summary>
    /// 是否允许刷新数据。
    /// </summary>
    bool EnableRefresh { get; set; }

    /// <summary>
    /// 刷新数据时的单元格插入模式。
    /// </summary>
    XlCellInsertionMode RefreshStyle { get; set; }

    /// <summary>
    /// 是否发生获取行溢出（只读）。
    /// </summary>
    bool FetchedRowOverflow { get; }

    /// <summary>
    /// 是否允许编辑表格数据。
    /// </summary>
    bool EnableEditing { get; set; }

    /// <summary>
    /// 是否保留列信息（如列名、数据类型等）。
    /// </summary>
    bool PreserveColumnInfo { get; set; }

    /// <summary>
    /// 是否保留格式（如字体、颜色、边框等）。
    /// </summary>
    bool PreserveFormatting { get; set; }

    /// <summary>
    /// 是否自动调整列宽以适应内容。
    /// </summary>
    bool AdjustColumnWidth { get; set; }


    // ========== 关联对象 ==========

    /// <summary>
    /// 获取数据目标区域（刷新后数据写入的位置）。
    /// </summary>
    IExcelRange? Destination { get; }

    /// <summary>
    /// 获取当前结果数据区域（实际包含数据的范围）。
    /// </summary>
    IExcelRange? ResultRange { get; }

    /// <summary>
    /// 获取关联的 ListObject（即 Excel 表格对象）。
    /// </summary>
    IExcelListObject? ListObject { get; }

    /// <summary>
    /// 获取关联的工作簿连接对象（用于外部数据源）。
    /// </summary>
    IExcelWorkbookConnection? WorkbookConnection { get; }


    // ========== 操作方法 ==========

    /// <summary>
    /// 删除整个表格对象（保留单元格数据，仅移除表格结构和连接）。
    /// </summary>
    void Delete();

    /// <summary>
    /// 刷新表格数据（从外部源重新加载）。
    /// </summary>
    /// <returns>刷新是否成功</returns>
    bool Refresh();
}