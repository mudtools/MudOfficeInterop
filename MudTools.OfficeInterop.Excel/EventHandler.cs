//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel;


/// <summary>
/// WorkbookNew事件处理委托
/// </summary>
/// <param name="workbook">新建工作簿</param>
public delegate void WorkbookNewEventHandler(IExcelWorkbook workbook);


/// <summary>
/// WorkbookOpen事件处理委托
/// </summary>
/// <param name="workbook">打开的工作簿</param>
public delegate void WorkbookOpenEventHandler(IExcelWorkbook workbook);

/// <summary>
/// WorkbookActivate事件处理委托
/// </summary>
/// <param name="workbook">激活的工作簿</param>
public delegate void WorkbookActivateEventHandler(IExcelWorkbook workbook);

/// <summary>
/// WorkbookDeactivate事件处理委托
/// </summary>
/// <param name="workbook">取消激活的工作簿</param>
public delegate void WorkbookDeactivateEventHandler(IExcelWorkbook workbook);

/// <summary>
/// WorkbookBeforeClose事件处理委托
/// </summary>
/// <param name="workbook">即将关闭的工作簿</param>
/// <param name="cancel">是否取消关闭操作</param>
public delegate void WorkbookBeforeCloseEventHandler(IExcelWorkbook workbook, ref bool cancel);

/// <summary>
/// WorkbookBeforeSave事件处理委托
/// </summary>
/// <param name="workbook">即将保存的工作簿</param>
/// <param name="saveAsUI">如果操作来自"另存为"对话框，则为true；如果操作来自"保存"或"另存为"命令，则为false</param>
/// <param name="Cancel">如果设置为true，则取消保存操作</param>
public delegate void WorkbookBeforeSaveEventHandler(IExcelWorkbook workbook, bool saveAsUI, ref bool Cancel);

/// <summary>
/// SheetChange事件处理委托
/// </summary>
/// <param name="sheet">发生变化的工作表</param>
/// <param name="target">发生变化的单元格区域</param>
public delegate void SheetChangeEventHandler(IExcelWorksheet sheet, IExcelRange target);

/// <summary>
/// SheetActivate事件处理委托
/// </summary>
/// <param name="sheet">激活的工作表</param>
public delegate void SheetActivateEventHandler(IExcelWorksheet sheet);

/// <summary>
/// 工作表删除前事件处理委托
/// </summary>
/// <param name="sheet">即将被删除的工作表</param>
public delegate void SheetBeforeDeleteEventHandler(IExcelWorksheet sheet);

/// <summary>
/// SheetDeactivate事件处理委托
/// </summary>
/// <param name="sheet">取消激活的工作表</param>
public delegate void SheetDeactivateEventHandler(IExcelWorksheet sheet);

/// <summary>
/// SheetSelectionChange事件处理委托
/// </summary>
/// <param name="sheet">工作表</param>
/// <param name="target">选中的区域</param>
public delegate void SheetSelectionChangeEventHandler(IExcelWorksheet sheet, IExcelRange target);

/// <summary>
/// SheetBeforeDoubleClick事件处理委托
/// </summary>
/// <param name="sheet">工作表</param>
/// <param name="target">双击的区域</param>
/// <param name="cancel">是否取消默认操作</param>
public delegate void SheetBeforeDoubleClickEventHandler(IExcelWorksheet sheet, IExcelRange target, ref bool cancel);

/// <summary>
/// SheetBeforeRightClick事件处理委托
/// </summary>
/// <param name="sheet">工作表</param>
/// <param name="target">右键单击的区域</param>
/// <param name="cancel">是否取消默认操作</param>
public delegate void SheetBeforeRightClickEventHandler(IExcelWorksheet sheet, IExcelRange target, ref bool cancel);

/// <summary>
/// SheetCalculate事件处理委托
/// </summary>
/// <param name="sheet">工作表</param>
public delegate void SheetCalculateEventHandler(IExcelWorksheet sheet);


/// <summary>
/// Change事件处理委托
/// </summary>
/// <param name="target">发生变化的单元格区域</param>
public delegate void ChangeEventHandler(IExcelRange target);

/// <summary>
/// SelectionChange事件处理委托
/// </summary>
/// <param name="target">选中的区域</param>
public delegate void SelectionChangeEventHandler(IExcelRange target);

/// <summary>
/// BeforeDoubleClick事件处理委托
/// </summary>
/// <param name="target">双击的区域</param>
/// <param name="cancel">是否取消默认操作</param>
public delegate void BeforeDoubleClickEventHandler(IExcelRange target, ref bool cancel);

/// <summary>
/// BeforeRightClick事件处理委托
/// </summary>
/// <param name="target">右键单击的区域</param>
/// <param name="cancel">是否取消默认操作</param>
public delegate void BeforeRightClickEventHandler(IExcelRange target, ref bool cancel);

/// <summary>
/// Activate事件处理委托
/// </summary>
public delegate void ActivateEventHandler();

/// <summary>
/// Deactivate事件处理委托
/// </summary>
public delegate void DeactivateEventHandler();

/// <summary>
/// BeforeDelete事件处理委托
/// </summary>
public delegate void BeforeDeleteEventHandler();

/// <summary>
/// 工作簿同步事件处理委托
/// </summary>
/// <param name="syncEventType">同步事件类型</param>
public delegate void WorkBookSyncEventHandler(MsoSyncEventType syncEventType);

/// <summary>
/// 工作簿工作表变更事件处理委托
/// </summary>
/// <param name="sh">工作表对象</param>
/// <param name="target">变更的目标单元格区域</param>
public delegate void WorkBookSheetChangeEventHandler(object sh, IExcelRange target);

/// <summary>
/// 工作簿保存后事件处理委托
/// </summary>
/// <param name="success">保存操作是否成功</param>
public delegate void WorkBookAfterSaveEventHandler(bool success);

/// <summary>
/// 工作簿新建图表事件处理委托
/// </summary>
/// <param name="chart">新建的图表对象</param>
public delegate void WorkBookNewChartEventHandler(IExcelChart chart);

/// <summary>
/// 工作簿打印前事件处理委托
/// </summary>
/// <param name="cancel">是否取消打印操作</param>
public delegate void WorkBookBeforePrintEventHandler(bool cancel);

/// <summary>
/// 工作簿行集完成前事件处理委托
/// </summary>
/// <param name="description">行集描述</param>
/// <param name="sheet">工作表名称</param>
/// <param name="success">操作是否成功</param>
public delegate void WorkBookBeforeRowsetCompleteEventHandler(string description, string sheet, bool success);

/// <summary>
/// 工作簿新建工作表事件处理委托
/// </summary>
/// <param name="worksheet">新建的工作表对象</param>
public delegate void WorkBookNewSheetEventHandler(IExcelWorksheet worksheet);

/// <summary>
/// 工作表数据透视表变更事件处理委托
/// </summary>
/// <param name="worksheet">工作表对象</param>
/// <param name="pivotTable">变更的数据透视表对象</param>
public delegate void WorkBookSheetPivotTableChangeEventHandler(IExcelWorksheet worksheet, IExcelPivotTable pivotTable);

/// <summary>
/// 数据透视表关闭连接事件处理委托
/// </summary>
/// <param name="pivotTable">数据透视表对象</param>
public delegate void WorkBookPivotTableCloseConnectionEventHandler(IExcelPivotTable pivotTable);

/// <summary>
/// 数据透视表打开连接事件处理委托
/// </summary>
/// <param name="pivotTable">数据透视表对象</param>
public delegate void WorkBookPivotTableOpenConnectionEventHandler(IExcelPivotTable pivotTable);

/// <summary>
/// PivotTableChangeSync事件处理委托
/// </summary>
/// <param name="excelPivotTable">Excel数据透视表对象</param>
public delegate void PivotTableChangeSyncEventHandler(IExcelPivotTable excelPivotTable);


/// <summary>
/// WindowResize事件处理委托
/// </summary>
/// <param name="Wb">Excel工作簿对象</param>
/// <param name="Wn">Excel窗口对象</param>
public delegate void WindowResizeEventHandler(IExcelWorkbook Wb, IExcelWindow Wn);

/// <summary>
/// WindowDeactivate事件处理委托
/// </summary>
/// <param name="Wb">Excel工作簿对象</param>
/// <param name="Wn">Excel窗口对象</param>
public delegate void WindowDeactivateEventHandler(IExcelWorkbook Wb, IExcelWindow Wn);

/// <summary>
/// WindowActivate事件处理委托
/// </summary>
/// <param name="Wb">Excel工作簿对象</param>
/// <param name="Wn">Excel窗口对象</param>
public delegate void WindowActivateEventHandler(IExcelWorkbook Wb, IExcelWindow Wn);

/// <summary>
/// Calculate事件处理委托
/// </summary>
public delegate void CalculateEventHandler();

/// <summary>
/// 图表激活事件：当图表被激活时触发
/// </summary>
/// <param name="chart">触发事件的 ExcelChart 实例</param>
public delegate void ChartActivateEventHandler(IExcelChart chart);

/// <summary>
/// 图表失活事件：当图表失去焦点时触发
/// </summary>
/// <param name="chart">触发事件的 ExcelChart 实例</param>
public delegate void ChartDeactivateEventHandler(IExcelChart chart);

/// <summary>
/// 图表选择事件：用户在图表上点击任意元素（数据点、图例、坐标轴等）时触发
/// </summary>
/// <param name="chart">触发事件的 ExcelChart 实例</param>
/// <param name="elementId">选中元素的类型，MsoChartElementType 枚举值（int）</param>
/// <param name="seriesIndex">系列索引（int），若选中的是系列或数据点，则为该系列编号；否则为 Missing.Value</param>
/// <param name="pointIndex">数据点索引（int），若选中的是单个数据点，则为该点编号；否则为 Missing.Value</param>
public delegate void ChartSelectEventHandler(IExcelChart chart, MsoChartElementType elementId, int seriesIndex, int pointIndex);

/// <summary>
/// 图表双击前事件：可取消默认双击行为（如打开格式对话框）
/// </summary>
/// <param name="cancel">设为 true 可取消默认行为</param>
/// <param name="elementId">选中元素的类型，MsoChartElementType 枚举值（int）</param>
/// <param name="seriesIndex">系列索引（int），若选中的是系列或数据点，则为该系列编号；否则为 Missing.Value</param>
/// <param name="pointIndex">数据点索引（int），若选中的是单个数据点，则为该点编号；否则为 Missing.Value</param>
public delegate void ChartBeforeDoubleClickEventHandler(MsoChartElementType elementId, int seriesIndex, int pointIndex, ref bool cancel);

/// <summary>
/// 图表右键单击前事件：可取消默认上下文菜单显示
/// </summary>
/// <param name="cancel">设为 true 可取消默认菜单</param>
public delegate void ChartBeforeRightClickEventHandler(ref bool cancel);

/// <summary>
/// 图表数据系列变化事件：当任一数据系列的数据发生变更时触发（如单元格值改变）
/// </summary>
/// <param name="chart">触发事件的 ExcelChart 实例</param>
/// <param name="seriesIndex">发生变化的系列索引（int）</param>
/// <param name="pointIndex">若仅单个点变化，则为该点索引；否则为 Missing.Value</param>
public delegate void ChartSeriesChangeEventHandler(IExcelChart chart, int seriesIndex, int pointIndex);

/// <summary>
/// 应用程序性能统计信息
/// </summary>
public class ApplicationPerformance
{
    public int MiscOperations { get; set; }

    /// <summary>
    /// 启动时间
    /// </summary>
    public DateTime StartupTime { get; set; }

    /// <summary>
    /// 运行时间（秒）
    /// </summary>
    public double RunTime { get; set; }

    /// <summary>
    /// 工作簿操作次数
    /// </summary>
    public int WorkbookOperations { get; set; }

    /// <summary>
    /// 计算操作次数
    /// </summary>
    public int CalculationOperations { get; set; }

    /// <summary>
    /// 文件操作次数
    /// </summary>
    public int FileOperations { get; set; }

    /// <summary>
    /// 宏执行次数
    /// </summary>
    public int MacroExecutions { get; set; }

    /// <summary>
    /// 平均响应时间（毫秒）
    /// </summary>
    public double AverageResponseTime { get; set; }

    /// <summary>
    /// 峰值内存使用（MB）
    /// </summary>
    public int PeakMemoryUsage { get; set; }
}

/// <summary>
/// 内存信息
/// </summary>
public class MemoryInfo
{
    /// <summary>
    /// 可用内存（MB）
    /// </summary>
    public int FreeMemory { get; set; }

    /// <summary>
    /// 总内存（MB）
    /// </summary>
    public int TotalMemory { get; set; }

    /// <summary>
    /// 已使用内存（MB）
    /// </summary>
    public int UsedMemory { get; set; }

    /// <summary>
    /// 内存使用率百分比
    /// </summary>
    public double UsagePercentage { get; set; }
}
