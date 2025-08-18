//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
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

public delegate void WorkbookBeforeSaveEventHandler(IExcelWorkbook workbook, bool SaveAsUI, ref bool Cancel);

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
/// Calculate事件处理委托
/// </summary>
public delegate void CalculateEventHandler();

public delegate void WindowResizeEventHandler(IExcelWorkbook Wb, IExcelWindow Wn);

public delegate void WindowDeactivateEventHandler(IExcelWorkbook Wb, IExcelWindow Wn);

public delegate void WindowActivateEventHandler(IExcelWorkbook Wb, IExcelWindow Wn);

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
