//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

using log4net;
using MudTools.OfficeInterop.Imps;
using System.Drawing;
using System.Globalization;

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// Excel Application 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.Application 对象的安全访问和资源管理
/// </summary>
internal class ExcelApplication : IExcelApplication
{
    /// <summary>
    /// 用于记录此类型运行时日志的 logger 实例。
    /// </summary>
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelApplication));

    /// <summary>
    /// 底层的 COM Application 对象
    /// </summary>
    private MsExcel.Application _application;

    private MsExcel.AppEvents_Event _appEvents_Event;

    /// <summary>
    /// 工作簿集合缓存
    /// </summary>
    private IExcelWorkbooks _workbooks;

    /// <summary>
    /// 错误检查选项缓存
    /// </summary>
    private IExcelErrorCheckingOptions _errorCheckingOptions;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    /// <summary>
    /// 启动时间
    /// </summary>
    private DateTime _startupTime;

    /// <summary>
    /// 性能统计信息
    /// </summary>
    private ApplicationPerformance _performanceStats;

    #region 事件字段

    private event WindowResizeEventHandler _windowResize;

    private event WindowDeactivateEventHandler _windowDeactivate;

    private event WindowActivateEventHandler _windowActivate;

    private event WorkbookNewEventHandler _newWorkbook;

    /// <summary>
    /// WorkbookOpen事件
    /// </summary>
    private event WorkbookOpenEventHandler _workbookOpen;

    /// <summary>
    /// WorkbookActivate事件
    /// </summary>
    private event WorkbookActivateEventHandler _workbookActivate;

    /// <summary>
    /// WorkbookDeactivate事件
    /// </summary>
    private event WorkbookDeactivateEventHandler _workbookDeactivate;

    /// <summary>
    /// WorkbookBeforeClose事件
    /// </summary>
    private event WorkbookBeforeCloseEventHandler _workbookBeforeClose;

    private event WorkbookBeforeSaveEventHandler _workbookBeforeSave;

    /// <summary>
    /// SheetChange事件
    /// </summary>
    private event SheetChangeEventHandler _sheetChange;

    /// <summary>
    /// SheetActivate事件
    /// </summary>
    private event SheetActivateEventHandler _sheetActivate;

    /// <summary>
    /// SheetDeactivate事件
    /// </summary>
    private event SheetDeactivateEventHandler _sheetDeactivate;

    /// <summary>
    /// SheetSelectionChange事件
    /// </summary>
    private event SheetSelectionChangeEventHandler _sheetSelectionChange;

    /// <summary>
    /// SheetBeforeDoubleClick事件
    /// </summary>
    private event SheetBeforeDoubleClickEventHandler _sheetBeforeDoubleClick;

    /// <summary>
    /// SheetBeforeRightClick事件
    /// </summary>
    private event SheetBeforeRightClickEventHandler _sheetBeforeRightClick;

    /// <summary>
    /// SheetCalculate事件
    /// </summary>
    private event SheetCalculateEventHandler _sheetCalculate;

    #endregion

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelApplication 实例（创建新的Excel应用程序）
    /// </summary>
    public ExcelApplication()
    {
        _application = new MsExcel.Application();
        _appEvents_Event = _application;
        _startupTime = DateTime.Now;
        _performanceStats = new ApplicationPerformance { StartupTime = _startupTime };
        _disposedValue = false;
        InitializeEvents();
    }

    /// <summary>
    /// 初始化 ExcelApplication 实例（使用现有的Excel应用程序）
    /// </summary>
    /// <param name="application">底层的 COM Application 对象</param>
    internal ExcelApplication(MsExcel.Application application)
    {
        _application = application ?? throw new ArgumentNullException(nameof(application));
        _appEvents_Event = _application;
        _startupTime = DateTime.Now;
        _performanceStats = new ApplicationPerformance { StartupTime = _startupTime };
        _disposedValue = false;
        InitializeEvents();
    }

    /// <summary>
    /// 初始化Excel事件处理
    /// </summary>
    private void InitializeEvents()
    {
        if (_appEvents_Event == null) return;

        _appEvents_Event.NewWorkbook += OnNewWorkbook;
        _appEvents_Event.WorkbookOpen += OnWorkbookOpen;
        _appEvents_Event.WorkbookActivate += OnWorkbookActivate;
        _appEvents_Event.WorkbookDeactivate += OnWorkbookDeactivate;
        _appEvents_Event.WorkbookBeforeClose += OnWorkbookBeforeClose;
        _appEvents_Event.WorkbookBeforeSave += OnWorkbookBeforeSave;
        _appEvents_Event.SheetChange += OnSheetChange;
        _appEvents_Event.SheetActivate += OnSheetActivate;
        _appEvents_Event.SheetDeactivate += OnSheetDeactivate;
        _appEvents_Event.SheetSelectionChange += OnSheetSelectionChange;
        _appEvents_Event.SheetBeforeDoubleClick += OnSheetBeforeDoubleClick;
        _appEvents_Event.SheetBeforeRightClick += OnSheetBeforeRightClick;
        _appEvents_Event.SheetCalculate += OnSheetCalculate;
        _appEvents_Event.WindowActivate += OnWindowActivate;
        _appEvents_Event.WindowDeactivate += OnWindowDeactivate;
        _appEvents_Event.WindowResize += OnWindowResize;
    }


    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">是否为显式释放</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放子对象
                _workbooks?.Dispose();
                _activeSheet?.Dispose();
                _errorCheckingOptions?.Dispose();
                _recentFiles?.Dispose();
                _sheets?.Dispose();
                _worksheets?.Dispose();
                _windows?.Dispose();
                _cells?.Dispose();
                _excelAddIns?.Dispose();
                // 移除事件处理
                DisConnectEvent();

                // 退出Excel应用程序（如果是我们创建的）
                if (_application != null)
                {
                    try
                    {
                        _application.Quit();
                    }
                    catch
                    {
                        // 忽略退出过程中的异常
                    }

                    Marshal.ReleaseComObject(_application);
                }
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _activeSheet = null;
            _cells = null;
            _recentFiles = null;
            _excelAddIns = null;
            _sheets = null;
            _worksheets = null;
            _application = null;
            _workbooks = null;
            _errorCheckingOptions = null;
        }

        _disposedValue = true;
    }

    private void DisConnectEvent()
    {
        if (_appEvents_Event == null)
            return;

        _appEvents_Event.NewWorkbook -= OnNewWorkbook;
        _appEvents_Event.WorkbookOpen -= OnWorkbookOpen;
        _appEvents_Event.WorkbookActivate -= OnWorkbookActivate;
        _appEvents_Event.WorkbookDeactivate -= OnWorkbookDeactivate;
        _appEvents_Event.WorkbookBeforeClose -= OnWorkbookBeforeClose;
        _appEvents_Event.SheetChange -= OnSheetChange;
        _appEvents_Event.SheetActivate -= OnSheetActivate;
        _appEvents_Event.SheetDeactivate -= OnSheetDeactivate;
        _appEvents_Event.SheetSelectionChange -= OnSheetSelectionChange;
        _appEvents_Event.SheetBeforeDoubleClick -= OnSheetBeforeDoubleClick;
        _appEvents_Event.SheetBeforeRightClick -= OnSheetBeforeRightClick;
        _appEvents_Event.SheetCalculate -= OnSheetCalculate;
    }

    ~ExcelApplication()
    {
        Dispose(false);
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取或设置单元格拖放功能是否启用
    /// </summary>
    public bool CellDragAndDrop
    {
        get => _application.CellDragAndDrop;
        set => _application.CellDragAndDrop = value;
    }


    /// <summary>
    /// 获取或设置是否启用实时预览
    /// </summary>
    public bool EnableLivePreview
    {
        get => _application.EnableLivePreview;
        set => _application.EnableLivePreview = value;
    }

    /// <summary>
    /// 获取或设置是否显示浮动工具栏
    /// </summary>
    public bool ShowSelectionFloaties
    {
        get => _application.ShowSelectionFloaties;
        set => _application.ShowSelectionFloaties = value;
    }

    /// <summary>
    /// 获取或设置是否显示开发工具选项卡
    /// </summary>
    public bool ShowDevTools
    {
        get => _application.ShowDevTools;
        set => _application.ShowDevTools = value;
    }

    /// <summary>
    /// 获取或设置是否忽略远程请求
    /// </summary>
    public bool IgnoreRemoteRequests
    {
        get => _application.IgnoreRemoteRequests;
        set => _application.IgnoreRemoteRequests = value;
    }

    /// <summary>
    /// 获取应用程序的名称
    /// </summary>
    public string Name => _application?.Name;

    /// <summary>
    /// 获取应用程序的版本号
    /// </summary>
    public string Version => _application?.Version;

    public int? Hwnd => _application?.Hwnd;

    public string? Build => _application?.Build.ToString();

    public IExcelCellFormat FindFormat
    {
        get
        {
            return _application != null ? new ExcelCellFormat(_application.FindFormat) : null;
        }
    }

    public XlCommentDisplayMode DisplayCommentIndicator
    {
        get => (XlCommentDisplayMode)_application.DisplayCommentIndicator;
        set => _application.DisplayCommentIndicator = (MsExcel.XlCommentDisplayMode)value;
    }

    public XlMousePointer Cursor
    {
        get => (XlMousePointer)_application.Cursor;
        set => _application.Cursor = (MsExcel.XlMousePointer)value;
    }
    public IOfficeLanguageSettings LanguageSettings => new OfficeLanguageSettings(_application.LanguageSettings);

    public IOfficeCommandBars CommandBars => new OfficeCommandBars(_application.CommandBars);

    /// <summary>
    /// 获取或设置应用程序是否可见
    /// </summary>
    public bool Visible
    {
        get => _application != null && _application.Visible;
        set
        {
            if (_application != null)
                _application.Visible = value;
        }
    }


    /// <summary>
    /// 获取或设置应用程序是否处于用户控制状态
    /// </summary>
    public bool UserControl
    {
        get => _application != null && _application.UserControl;
        set
        {
            if (_application != null)
                _application.UserControl = value;
        }
    }

    /// <summary>
    /// 获取或设置应用程序是否显示警告信息
    /// </summary>
    public bool DisplayAlerts
    {
        get => _application != null && _application.DisplayAlerts;
        set
        {
            if (_application != null)
                _application.DisplayAlerts = value;
        }
    }

    /// <summary>
    /// 获取或设置应用程序是否启用事件
    /// </summary>
    public bool EnableEvents
    {
        get => _application != null && _application.EnableEvents;
        set
        {
            if (_application != null)
                _application.EnableEvents = value;
        }
    }

    /// <summary>
    /// 获取或设置应用程序是否启用动画
    /// </summary>
    public bool EnableAnimations
    {
        get => _application != null && _application.EnableAnimations;
        set
        {
            if (_application != null)
                _application.EnableAnimations = value;
        }
    }

    public XlEnableCancelKey EnableCancelKey
    {
        get => (XlEnableCancelKey)_application.EnableCancelKey;
        set
        {
            if (_application != null)
                _application.EnableCancelKey = (MsExcel.XlEnableCancelKey)value;
        }
    }

    /// <summary>
    /// 获取应用程序的当前路径
    /// </summary>
    public string Path => _application?.Path;

    /// <summary>
    /// 获取或设置应用程序的当前用户
    /// </summary>
    public string UserName
    {
        get => _application?.UserName;
        set
        {
            if (_application != null && value != null)
                _application.UserName = value;
        }
    }

    /// <summary>
    /// 获取应用程序的组织名称
    /// </summary>
    public string OrganizationName => _application?.OrganizationName?.ToString();


    /// <summary>
    /// 获取或设置应用程序窗口的标题栏文本。
    /// </summary>
    public string Caption
    {
        get => _application?.Caption?.ToString();
        set
        {
            if (_application != null)
                _application.Caption = value;
        }
    }

    /// <summary>
    /// 获取或设置新建工作簿中包含的工作表数量。
    /// </summary>
    public int SheetsInNewWorkbook
    {
        get => _application?.SheetsInNewWorkbook ?? 3; // 默认值通常为3
        set
        {
            if (_application != null)
                _application.SheetsInNewWorkbook = value;
        }
    }

    /// <summary>
    /// 获取或设置应用程序的默认字体名称。
    /// </summary>
    public string StandardFont
    {
        get => _application?.StandardFont?.ToString();
        set
        {
            if (_application != null)
                _application.StandardFont = value;
        }
    }

    /// <summary>
    /// 获取或设置应用程序的默认字体大小。
    /// </summary>
    public double StandardFontSize
    {
        get => _application?.StandardFontSize ?? 10; // 默认值示例
        set
        {
            if (_application != null)
                _application.StandardFontSize = value;
        }
    }

    /// <summary>
    /// 获取或设置是否启用迭代计算。
    /// </summary>
    public bool Iteration
    {
        get => _application?.Iteration ?? false;
        set
        {
            if (_application != null)
                _application.Iteration = value;
        }
    }

    /// <summary>
    /// 获取或设置迭代计算时的最大误差。
    /// </summary>
    public double MaxChange
    {
        get => _application?.MaxChange ?? 0.001; // 默认值示例
        set
        {
            if (_application != null)
                _application.MaxChange = value;
        }
    }

    /// <summary>
    /// 获取或设置迭代计算时的最大迭代次数。
    /// </summary>
    public int MaxIterations
    {
        get => _application?.MaxIterations ?? 100; // 默认值示例
        set
        {
            if (_application != null)
                _application.MaxIterations = value;
        }
    }

    /// <summary>
    /// 获取或设置是否在状态栏显示计算过程。
    /// </summary>
    public bool DisplayInsertOptions
    {
        get => _application?.DisplayInsertOptions ?? true;
        set
        {
            if (_application != null)
                _application.DisplayInsertOptions = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示最近使用的文件列表。
    /// </summary>
    public bool DisplayRecentFiles
    {
        get => _application?.DisplayRecentFiles ?? true;
        set
        {
            if (_application != null)
                _application.DisplayRecentFiles = value;
        }
    }


    /// <summary>
    /// 获取或设置是否允许在单元格内直接编辑。
    /// </summary>
    public bool EditDirectlyInCell
    {
        get => _application?.EditDirectlyInCell ?? true;
        set
        {
            if (_application != null)
                _application.EditDirectlyInCell = value;
        }
    }

    /// <summary>
    /// 获取或设置是否启用自动完成功能。
    /// </summary>
    public bool EnableAutoComplete
    {
        get => _application?.EnableAutoComplete ?? true;
        set
        {
            if (_application != null)
                _application.EnableAutoComplete = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示图表工具提示中的系列名称。
    /// </summary>
    public bool ShowChartTipNames
    {
        get => _application?.ShowChartTipNames ?? true;
        set
        {
            if (_application != null)
                _application.ShowChartTipNames = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示图表工具提示中的数值。
    /// </summary>
    public bool ShowChartTipValues
    {
        get => _application?.ShowChartTipValues ?? true;
        set
        {
            if (_application != null)
                _application.ShowChartTipValues = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示屏幕提示。
    /// </summary>
    public bool ShowToolTips
    {
        get => _application?.ShowToolTips ?? true;
        set
        {
            if (_application != null)
                _application.ShowToolTips = value;
        }
    }

    /// <summary>
    /// 获取或设置打开文件时是否提示更新链接。
    /// </summary>
    public bool AskToUpdateLinks
    {
        get => _application?.AskToUpdateLinks ?? true;
        set
        {
            if (_application != null)
                _application.AskToUpdateLinks = value;
        }
    }

    /// <summary>
    /// 获取或设置覆盖文件时是否显示警告。
    /// </summary>
    public bool AlertBeforeOverwriting
    {
        get => _application?.AlertBeforeOverwriting ?? true;
        set
        {
            if (_application != null)
                _application.AlertBeforeOverwriting = value;
        }
    }

    /// <summary>
    /// 获取或设置过渡菜单键。
    /// </summary>
    public string TransitionMenuKey
    {
        get => _application?.TransitionMenuKey?.ToString();
        set
        {
            if (_application != null)
                _application.TransitionMenuKey = value;
        }
    }

    /// <summary>
    /// 获取或设置过渡菜单键的操作。
    /// </summary>
    public int TransitionMenuKeyAction
    {
        get => _application?.TransitionMenuKeyAction ?? 0;
        set
        {
            if (_application != null)
                _application.TransitionMenuKeyAction = value;
        }
    }

    /// <summary>
    /// 获取或设置按 Enter 键后是否移动选定区域。
    /// </summary>
    public bool MoveAfterReturn
    {
        get => _application?.MoveAfterReturn ?? true;
        set
        {
            if (_application != null)
                _application.MoveAfterReturn = value;
        }
    }

    /// <summary>
    /// 获取或设置按 Enter 键后移动选定区域的方向。
    /// </summary>
    public XlDirection MoveAfterReturnDirection
    {
        get => (XlDirection)_application?.MoveAfterReturnDirection;
        set
        {
            if (_application != null)
                _application.MoveAfterReturnDirection = (MsExcel.XlDirection)value;
        }
    }

    /// <summary>
    /// 获取应用程序窗口的可用高度（像素）。
    /// </summary>
    public double UsableHeight
    {
        get => _application?.UsableHeight ?? 0;
    }

    /// <summary>
    /// 获取应用程序窗口的可用宽度（像素）。
    /// </summary>
    public double UsableWidth
    {
        get => _application?.UsableWidth ?? 0;
    }

    /// <summary>
    /// 获取或设置状态栏显示的文本。
    /// </summary>
    public string StatusBar
    {
        get => _application?.StatusBar?.ToString();
        set
        {
            if (_application != null)
                _application.StatusBar = value; // 设置为 null 可恢复默认状态栏
        }
    }


    #endregion

    #region 工作簿管理

    /// <summary>
    /// 获取应用程序中的工作簿集合
    /// </summary>
    public IExcelWorkbooks Workbooks => _workbooks ??= new ExcelWorkbooks(_application?.Workbooks);

    /// <summary>
    /// 获取当前活动的工作簿
    /// </summary>
    public IExcelWorkbook? ActiveWorkbook
    {
        get
        {
            var workbook = _application?.ActiveWorkbook;
            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
    }

    public IExcelWindow? ActiveWindow
    {
        get
        {
            var window = _application?.ActiveWindow;
            return window != null ? new ExcelWindow(window, ActiveWorkbook) : null;
        }
    }

    public IExcelWorksheetFunction WorksheetFunction
    {
        get
        {
            var worksheetFunction = _application?.WorksheetFunction;
            return worksheetFunction != null ? new ExcelWorksheetFunction(worksheetFunction) : null;
        }
    }

    /// <summary>
    /// 获取应用程序的ThisWorkbook
    /// </summary>
    public IExcelWorkbook ThisWorkbook
    {
        get
        {
            var workbook = _application?.ThisWorkbook as MsExcel.Workbook;
            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
    }

    /// <summary>
    /// 获取工作簿的数量
    /// </summary>
    public int WorkbooksCount => _application?.Workbooks?.Count ?? 0;

    /// <summary>
    /// 获取指定索引的工作簿
    /// </summary>
    /// <param name="index">工作簿索引</param>
    /// <returns>工作簿对象</returns>
    public IExcelWorkbook GetWorkbook(int index)
    {
        if (_application?.Workbooks == null || index < 1 || index > WorkbooksCount)
            return null;

        try
        {
            var workbook = _application.Workbooks[index] as MsExcel.Workbook;
            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 获取指定名称的工作簿
    /// </summary>
    /// <param name="name">工作簿名称</param>
    /// <returns>工作簿对象</returns>
    public IExcelWorkbook GetWorkbook(string name)
    {
        if (_application?.Workbooks == null || string.IsNullOrEmpty(name))
            return null;

        try
        {
            var workbook = _application.Workbooks[name] as MsExcel.Workbook;
            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region 工作表管理
    public IExcelWorksheet? ActiveSheetWarp
    {
        get
        {
            var obj = ActiveSheet;
            if (obj is IExcelWorksheet worksheet)
            {
                return worksheet;
            }
            return null;
        }
    }

    private ICommonWorksheet? _activeSheet;

    /// <summary>
    /// 获取当前活动的工作表
    /// </summary>
    public ICommonWorksheet? ActiveSheet
    {
        get
        {
            if (_activeSheet != null)
                return _activeSheet;
            if (_application?.ActiveSheet == null)
                return null;

            object activeSheet = _application.ActiveSheet;

            if (activeSheet is MsExcel.Worksheet)
            {
                MsExcel.Worksheet ws = (MsExcel.Worksheet)activeSheet;
                _activeSheet = new ExcelWorksheet(ws);
            }
            else if (activeSheet is MsExcel.Chart)
            {
                MsExcel.Chart chart = (MsExcel.Chart)activeSheet;
                _activeSheet = new ExcelChart(chart);
            }
            else if (activeSheet is MsExcel.ChartObject)
            {
                MsExcel.ChartObject chartObj = (MsExcel.ChartObject)activeSheet;
                _activeSheet = new ExcelChartObject(chartObj);
            }
            else if (activeSheet is MsExcel.DialogSheet)
            {
                MsExcel.DialogSheet dialog = (MsExcel.DialogSheet)activeSheet;
            }
            return _activeSheet;
        }
    }

    /// <summary>
    /// 获取当前活动的单元格区域
    /// </summary>
    public IExcelRange? ActiveCell
    {
        get
        {
            if (_application?.ActiveCell == null)
                return null;
            MsExcel.Range? range = _application?.ActiveCell as MsExcel.Range;
            return range != null ? new ExcelRange(range) : null;
        }
    }

    /// <summary>
    /// 表示用户当前在Excel界面中选中的对象，并且在绝大多数情况下，它返回的是一个Range对象，用于操作工作表中的单元格或单元格区域。
    /// </summary>
    public object? Selection => Utils.CreateSelectionType(_application.Selection);



    /// <summary>
    /// 表示用户当前在Excel界面中选中Range对象，用于操作工作表中的单元格或单元格区域。
    /// </summary>
    public IExcelRange? SelectionRang
    {
        get
        {
            if (_application.Selection is MsExcel.Range rang)
                return new ExcelRange(rang);
            return null;
        }
    }

    private IExcelRows _activeRows;
    /// <summary>
    /// 获取当前活动行集合
    /// </summary>
    public IExcelRows ActiveRows
    {
        get
        {
            if (_activeRows != null) return _activeRows;

            if (_application.Selection is MsExcel.Range range)
            {
                _activeRows = new ExcelRange(range.Rows);
            }
            return _activeRows;
        }
    }

    public IExcelRange Columns
    {
        get
        {
            if (_application?.Columns != null) ;
            return new ExcelRange(_application.Columns);
        }
    }

    public IExcelRange Rows
    {
        get
        {
            if (_application?.Rows != null) ;
            return new ExcelRange(_application.Rows);
        }
    }

    /// <summary>
    /// 获取当前活动列集合
    /// </summary>
    public IExcelColumns ActiveColumns
    {
        get
        {
            if (_application.Selection is MsExcel.Range range)
            {
                return new ExcelRange(range.Columns);
            }
            return null;
        }
    }
    private IExcelRange _cells;
    public IExcelRange Cells => _cells ??= new ExcelRange(_application?.Cells);

    private IExcelSheets _sheets;
    public IExcelSheets Sheets => _sheets ??= new ExcelSheets(_application?.Sheets);

    private IExcelSheets _worksheets;
    public IExcelSheets Worksheets => _worksheets ??= new ExcelSheets(_application?.Worksheets);

    private IExcelRecentFiles _recentFiles;
    public IExcelRecentFiles RecentFiles => _recentFiles ??= new ExcelRecentFiles(_application?.RecentFiles);

    private IExcelWindows _windows;

    public IExcelWindows Windows => _windows ??= new ExcelWindows(_application?.Windows, this);

    private IExcelProtectedViewWindows _protectedViewWindows;
    public IExcelProtectedViewWindows ProtectedViewWindows => _protectedViewWindows ??= new ExcelProtectedViewWindows(_application?.ProtectedViewWindows);

    public IExcelProtectedViewWindow? ActiveProtectedViewWindow
    {
        get
        {
            var window = _application?.ActiveProtectedViewWindow;
            return window != null ? new ExcelProtectedViewWindow(window) : null;
        }
    }


    private IExcelAddIns? _excelAddIns;

    public IExcelAddIns? AddIns => _excelAddIns ??= new ExcelAddIns(_application?.AddIns);

    #endregion

    #region 计算设置

    /// <summary>
    /// 获取或设置计算模式
    /// </summary>
    public XlCalculation Calculation
    {
        get => _application != null ? (XlCalculation)_application.Calculation : XlCalculation.xlCalculationManual;
        set
        {
            if (_application != null)
                _application.Calculation = (MsExcel.XlCalculation)value;
        }
    }

    /// <summary>
    /// 获取或设置是否自动重算
    /// </summary>
    public bool CalculateBeforeSave
    {
        get => _application != null && _application.CalculateBeforeSave;
        set
        {
            if (_application != null)
                _application.CalculateBeforeSave = value;
        }
    }

    /// <summary>
    /// 获取或设置是否启用多线程计算
    /// </summary>
    public bool MultiThreadedCalculation
    {
        get => _application != null && _application.MultiThreadedCalculation.Enabled;
        set
        {
            if (_application != null)
                _application.MultiThreadedCalculation.Enabled = value;
        }
    }

    /// <summary>
    /// 手动计算所有打开的工作簿
    /// </summary>
    public void Calculate()
    {
        _application?.Calculate();
        _performanceStats.CalculationOperations++;
    }

    /// <summary>
    /// 重新计算所有打开的工作簿
    /// </summary>
    public void CalculateFull()
    {
        _application?.CalculateFull();
        _performanceStats.CalculationOperations++;
    }

    /// <summary>
    /// 计算指定工作表
    /// </summary>
    /// <param name="worksheet">要计算的工作表</param>
    public void CalculateWorksheet(IExcelWorksheet worksheet)
    {
        if (_application == null || worksheet == null) return;

        try
        {
            var excelWorksheet = worksheet as ExcelWorksheet;
            excelWorksheet?.Worksheet?.Calculate();
            _performanceStats.CalculationOperations++;
        }
        catch
        {
            // 忽略计算过程中的异常
        }
    }

    /// <summary>
    /// 获取多个区域的交集区域
    /// </summary>
    /// <param name="ranges">要计算交集的区域集合</param>
    /// <returns>交集区域（无交集时返回 null）</returns>
    /// <exception cref="ArgumentNullException">输入区域为空时抛出</exception>
    /// <exception cref="ArgumentException">区域数量不足时抛出</exception>
    public IExcelRange Intersect(params IExcelRange[] ranges)
    {
        // 参数验证
        if (ranges == null)
            throw new ArgumentNullException(nameof(ranges));
        if (ranges.Length < 2)
            throw new ArgumentException("至少需要两个区域进行交集计算", nameof(ranges));

        try
        {
            // 将自定义范围对象转换为原生 Range 对象数组
            var nativeRanges = ranges
                .Where(r => r != null)
                .Select(r => (r as ExcelRange)?.InternalRange)
                .Where(r => r != null)
                .ToArray();

            if (nativeRanges.Length < 2) return null;

            // 检查所有区域是否在同一个工作表
            var firstSheet = nativeRanges[0].Worksheet;
            if (nativeRanges.Any(r => r.Worksheet.Name != firstSheet.Name))
            {
                throw new InvalidOperationException("所有区域必须在同一个工作表内");
            }

            // 调用原生 Intersect 方法
            MsExcel.Range resultRange = _application.Intersect(nativeRanges[0], nativeRanges[1]);

            // 处理多个区域的情况
            for (int i = 2; i < nativeRanges.Length && resultRange != null; i++)
            {
                resultRange = _application.Intersect(resultRange, nativeRanges[i]);
            }

            return resultRange != null ? new ExcelRange(resultRange) : null;
        }
        catch (COMException ex) when (ex.ErrorCode == -2146827284) // 0x800A03EC
        {
            // 处理无交集的特殊情况 (Excel 错误代码)
            return null;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("计算区域交集失败", ex);
        }
    }

    /// <summary>
    /// 获取多个区域的并集区域
    /// </summary>
    /// <param name="ranges">要合并的区域集合</param>
    /// <returns>并集区域（无效输入时返回 null）</returns>
    /// <exception cref="ArgumentNullException">输入区域为空时抛出</exception>
    /// <exception cref="ArgumentException">区域数量不足时抛出</exception>
    public IExcelRange Union(params IExcelRange[] ranges)
    {
        // 参数验证
        if (ranges == null)
            throw new ArgumentNullException(nameof(ranges));
        if (ranges.Length < 1)
            throw new ArgumentException("至少需要一个区域进行合并", nameof(ranges));

        try
        {
            // 处理单个区域的情况
            if (ranges.Length == 1)
                return ranges[0];

            // 将自定义范围对象转换为原生 Range 对象数组
            MsExcel.Range?[] nativeRanges = ranges
                  .Where(r => r != null)
                  .Select(r => (r as ExcelRange)?.InternalRange)
                  .Where(r => r != null)
                  .ToArray();

            if (nativeRanges.Length == 0) return null;

            // 检查所有区域是否在同一个工作表
            MsExcel.Worksheet firstSheet = nativeRanges[0].Worksheet;
            if (nativeRanges.Any(r => r.Worksheet.Name != firstSheet.Name))
            {
                throw new InvalidOperationException("所有区域必须在同一个工作表内");
            }

            // 初始化为第一个区域
            MsExcel.Range resultRange = nativeRanges[0];

            // 迭代合并所有区域
            for (int i = 1; i < nativeRanges.Length; i++)
            {
                resultRange = _application.Union(resultRange, nativeRanges[i]);
            }

            return resultRange != null ? new ExcelRange(resultRange) : null;
        }
        catch (COMException ex)
        {
            // 处理常见的 COM 异常
            const int RPC_E_SERVERFAULT = unchecked((int)0x80010105);
            if (ex.ErrorCode == RPC_E_SERVERFAULT)
            {
                throw new InvalidOperationException("区域合并操作超时，请尝试减小区域范围", ex);
            }
            throw new InvalidOperationException("计算区域并集失败", ex);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("计算区域并集时发生错误", ex);
        }
    }

    /// <summary>
    /// 检查多个区域是否相邻（可形成连续区域）
    /// </summary>
    public bool AreContiguous(params IExcelRange[] ranges)
    {
        try
        {
            var unionRange = Union(ranges);
            if (unionRange == null) return false;

            // 连续区域应该有相同的行数和列数
            return unionRange.Areas.Count == 1;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 合并区域并应用格式
    /// </summary>
    public void FormatUnionRange(IExcelCellFormat format, params IExcelRange[] ranges)
    {
        var unionRange = Union(ranges);
        if (unionRange != null)
        {
            ApplyCellFormat(unionRange, format);
        }
    }

    /// <summary>
    /// 将单元格格式应用到指定区域
    /// </summary>
    /// <param name="range">目标区域</param>
    /// <param name="format">单元格格式配置</param>
    /// <param name="applyToSubAreas">是否应用到不连续的子区域</param>
    public void ApplyCellFormat(IExcelRange range, IExcelCellFormat format, bool applyToSubAreas = false)
    {
        if (range == null || format == null) return;

        try
        {
            // 获取底层原生Range对象
            var nativeRange = (range as ExcelRange)?.InternalRange;
            if (nativeRange == null) return;

            // 处理不连续区域
            if (applyToSubAreas && nativeRange.Areas.Count > 1)
            {
                foreach (MsExcel.Range area in nativeRange.Areas)
                {
                    ApplyFormatToRange(area, format);
                }
            }
            else
            {
                ApplyFormatToRange(nativeRange, format);
            }
        }
        catch (COMException ex)
        {
            HandleFormattingError(ex, range, format);
        }
    }

    /// <summary>
    /// 实际应用格式到原生Range对象
    /// </summary>
    private void ApplyFormatToRange(MsExcel.Range range, IExcelCellFormat format)
    {
        // 1. 字体格式
        if (format.Font != null)
        {
            range.Font.Bold = format.Font.Bold;
            range.Font.Italic = format.Font.Italic;
            range.Font.Underline = (MsExcel.XlUnderlineStyle)format.Font.Underline;
            range.Font.Size = format.Font.Size;
            range.Font.Color = ColorTranslator.ToOle(format.Font.Color);
            range.Font.Name = format.Font.Name;
        }

        // 3. 边框设置
        if (format.Borders != null)
        {
            ApplyBorders(range.Borders, format.Borders);
        }

        // 4. 对齐方式
        range.HorizontalAlignment = (MsExcel.XlHAlign)format.HorizontalAlignment;
        range.VerticalAlignment = (MsExcel.XlVAlign)format.VerticalAlignment;
        range.WrapText = format.WrapText;
        range.Orientation = format.Orientation;

        // 5. 数字格式
        if (!string.IsNullOrEmpty(format.NumberFormat?.ToString()))
        {
            range.NumberFormat = format.NumberFormat;
        }
    }

    /// <summary>
    /// 应用边框格式
    /// </summary>
    private void ApplyBorders(MsExcel.Borders borders, IExcelBorders borderFormat)
    {
        // 定义边框位置映射
        var borderTypes = new Dictionary<XlBordersIndex, MsExcel.XlBordersIndex>
    {
        { XlBordersIndex.xlEdgeLeft, MsExcel.XlBordersIndex.xlEdgeLeft },
        { XlBordersIndex.xlEdgeTop, MsExcel.XlBordersIndex.xlEdgeTop },
        { XlBordersIndex.xlEdgeRight, MsExcel.XlBordersIndex.xlEdgeRight },
        { XlBordersIndex.xlEdgeBottom, MsExcel.XlBordersIndex.xlEdgeBottom },
        { XlBordersIndex.xlInsideHorizontal, MsExcel.XlBordersIndex.xlInsideHorizontal },
        { XlBordersIndex.xlInsideVertical, MsExcel.XlBordersIndex.xlInsideVertical },
        { XlBordersIndex.xlDiagonalDown, MsExcel.XlBordersIndex.xlDiagonalDown },
        { XlBordersIndex.xlDiagonalUp, MsExcel.XlBordersIndex.xlDiagonalUp }
    };

        // 应用全局边框设置
        if (borderFormat.ApplyToAll)
        {
            borders.LineStyle = (MsExcel.XlLineStyle)borderFormat.LineStyle;
            borders.Weight = (MsExcel.XlBorderWeight)borderFormat.Weight;
            borders.Color = ColorTranslator.ToOle(borderFormat.Color);
            return;
        }

        // 应用特定位置边框
        foreach (var position in borderFormat.CustomBorders)
        {
            if (borderTypes.TryGetValue(position.Key, out var borderIndex))
            {
                var border = borders[borderIndex];
                border.LineStyle = (MsExcel.XlLineStyle)borderFormat.LineStyle;
                border.Weight = (MsExcel.XlBorderWeight)borderFormat.Weight;
                border.Color = ColorTranslator.ToOle(borderFormat.Color);
            }
        }
    }

    /// <summary>
    /// 处理格式应用错误
    /// </summary>
    private void HandleFormattingError(COMException ex, IExcelRange range, IExcelCellFormat format)
    {
        const int FORMAT_PROTECTED = unchecked((int)0x800A03EC); // Excel 错误代码

        if (ex.ErrorCode == FORMAT_PROTECTED)
        {
            // 尝试临时解除保护
            var worksheet = range.Worksheet;
            bool wasProtected = worksheet.IsProtected;
            string password = null; // 实际应用中可能需要存储密码

            try
            {
                if (wasProtected)
                {
                    worksheet.Unprotect(password);
                }

                ApplyCellFormat(range, format);
            }
            finally
            {
                if (wasProtected)
                {
                    worksheet.Protect(password);
                }
            }
        }
        else
        {
            throw new Exception($"无法应用格式到区域 {range.Address}", ex);
        }
    }


    /// <summary>
    /// 计算 Excel 公式
    /// </summary>
    /// <param name="formula">Excel 公式字符串</param>
    /// <returns>计算结果</returns>
    public object Evaluate(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
            throw new ArgumentException("Formula cannot be null or empty.", nameof(formula));

        try
        {
            object result = _application.Evaluate(formula);
            return ProcessResult(result);
        }
        catch (COMException ex)
        {
            throw new ExcelOperationException($"Error evaluating formula: {formula}", ex);
        }
    }

    private object ProcessResult(object result)
    {
        // 处理错误值
        if (result is MsExcel.XlCVError errorValue && errorValue != MsExcel.XlCVError.xlErrNull)
        {
            throw new ExcelOperationException($"Excel evaluation error: {GetErrorDescription(errorValue)}");
        }

        // 处理日期值（Excel日期存储为双精度浮点数）
        if (result is double d && d.IsExcelDateSerial())
        {
            return DateTime.FromOADate(d);
        }

        return result;
    }


    private string GetErrorDescription(MsExcel.XlCVError error)
    {
        return error switch
        {
            MsExcel.XlCVError.xlErrDiv0 => "#DIV/0! - Division by zero",
            MsExcel.XlCVError.xlErrNA => "#N/A - Value not available",
            MsExcel.XlCVError.xlErrName => "#NAME? - Invalid name",
            MsExcel.XlCVError.xlErrNull => "#NULL! - Null reference",
            MsExcel.XlCVError.xlErrNum => "#NUM! - Invalid number",
            MsExcel.XlCVError.xlErrRef => "#REF! - Invalid reference",
            MsExcel.XlCVError.xlErrValue => "#VALUE! - Invalid value",
            _ => $"Unknown error: {error}"
        };
    }


    /// <summary>
    /// 计算 Excel 公式（带参数）
    /// </summary>
    /// <param name="formula">包含占位符的公式模板</param>
    /// <param name="args">公式参数</param>
    /// <returns>计算结果</returns>
    public object Evaluate(string formula, params object[] args)
    {
        if (string.IsNullOrWhiteSpace(formula))
            throw new ArgumentException("Formula cannot be null or empty.", nameof(formula));

        string formattedFormula = FormatFormula(formula, args);
        return Evaluate(formattedFormula);
    }

    private string FormatFormula(string formula, object[] args)
    {
        for (int i = 0; i < args.Length; i++)
        {
            string placeholder = $"{{{i}}}";
            string replacement = FormatArgument(args[i]);

            formula = formula.Replace(placeholder, replacement);
        }
        return formula;
    }

    private string FormatArgument(object arg)
    {
        return arg switch
        {
            null => "\"\"",
            string s => $"\"{s.Replace("\"", "\"\"")}\"", // 处理字符串中的引号
            bool b => b ? "TRUE" : "FALSE",
            DateTime dt => dt.ToOADate().ToString(CultureInfo.InvariantCulture),
            double d => d.ToString(CultureInfo.InvariantCulture),
            int i => i.ToString(),
            _ => arg.ToString()
        };
    }

    /// <summary>
    /// 强类型计算结果（数值类型）
    /// </summary>
    public double EvaluateToNumber(string formula)
    {
        object result = Evaluate(formula);
        return result.ConvertToDouble();
    }

    /// <summary>
    /// 强类型计算结果（布尔类型）
    /// </summary>
    public bool EvaluateToBool(string formula)
    {
        object result = Evaluate(formula);
        return result.ConvertToBool();
    }

    /// <summary>
    /// 强类型计算结果（日期类型）
    /// </summary>
    public DateTime EvaluateToDateTime(string formula)
    {
        object result = Evaluate(formula);
        return result.ConvertToDateTime();
    }

    /// <summary>
    /// 强类型计算结果（字符串类型）
    /// </summary>
    public string EvaluateToString(string formula)
    {
        object result = Evaluate(formula);
        return result?.ToString() ?? string.Empty;
    }

    /// <summary>
    /// 计算并返回二维数组结果（用于范围计算结果）
    /// </summary>
    public object[,] EvaluateToArray(string formula)
    {
        object result = Evaluate(formula);
        return result.ConvertToArray();
    }

    #endregion

    #region 屏幕和显示

    public void Activate()
    {
        try
        {
            _application.Visible = true;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to activate Word application.", ex);
        }
    }

    public void Quit()
    {
        try
        {
            _application.Quit();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to quit Word application.", ex);
        }
    }

    public IOfficeFileDialog CreateFileDialog(MsoFileDialogType fileDialogType)
    {
        MsCore.FileDialog dialog = _application.FileDialog[(MsCore.MsoFileDialogType)fileDialogType];
        return new OfficeFileDialog(dialog);
    }

    /// <summary>
    /// 获取或设置是否显示滚动条
    /// </summary>
    public bool DisplayScrollBars
    {
        get => _application != null && _application.DisplayScrollBars;
        set
        {
            if (_application != null)
                _application.DisplayScrollBars = value;
        }
    }

    public bool DisplayFullScreen
    {
        get => _application != null && _application.DisplayFullScreen;
        set
        {
            if (_application != null)
                _application.DisplayFullScreen = value;
        }
    }

    public bool ShowWindowsInTaskbar
    {
        get => _application != null && _application.ShowWindowsInTaskbar;
        set
        {
            if (_application != null)
                _application.ShowWindowsInTaskbar = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示公式栏
    /// </summary>
    public bool DisplayFormulaBar
    {
        get => _application != null && _application.DisplayFormulaBar;
        set
        {
            if (_application != null)
                _application.DisplayFormulaBar = value;
        }
    }

    /// <summary>
    /// 获取或设置屏幕更新是否启用
    /// </summary>
    public bool ScreenUpdating
    {
        get => _application != null && _application.ScreenUpdating;
        set
        {
            if (_application != null)
                _application.ScreenUpdating = value;
        }
    }

    /// <summary>
    /// 获取或设置是否显示状态栏
    /// </summary>
    public bool DisplayStatusBar
    {
        get => _application != null && _application.DisplayStatusBar;
        set
        {
            if (_application != null)
                _application.DisplayStatusBar = value;
        }
    }

    /// <summary>
    /// 获取或设置窗口状态
    /// </summary>
    public int WindowState
    {
        get => _application != null ? Convert.ToInt32(_application.WindowState) : 0;
        set
        {
            if (_application != null)
                _application.WindowState = (MsExcel.XlWindowState)value;
        }
    }

    /// <summary>
    /// 获取或设置应用程序的高度
    /// </summary>
    public float Height
    {
        get => Convert.ToSingle(_application?.Height);
        set
        {
            if (_application != null)
                _application.Height = value;
        }
    }

    /// <summary>
    /// 获取或设置应用程序的宽度
    /// </summary>
    public float Width
    {
        get => Convert.ToSingle(_application?.Width);
        set
        {
            if (_application != null)
                _application.Width = value;
        }
    }

    /// <summary>
    /// 获取或设置应用程序的左边距
    /// </summary>
    public float Left
    {
        get => Convert.ToSingle(_application?.Left);
        set
        {
            if (_application != null)
                _application.Left = value;
        }
    }

    /// <summary>
    /// 获取或设置应用程序的顶边距
    /// </summary>
    public float Top
    {
        get => Convert.ToSingle(_application?.Top);
        set
        {
            if (_application != null)
                _application.Top = value;
        }
    }

    #endregion

    #region 文件操作

    public IOfficeFileDialog FileDialog(MsoFileDialogType fileDialogType) => new OfficeFileDialog(_application.FileDialog[(MsCore.MsoFileDialogType)fileDialogType]);

    public IExcelWorkbook BlankWorkbook()
    {
        try
        {
            var book = new ExcelWorkbook(
                _application.Workbooks.Add());
            return book;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create blank workbook.", ex);
        }
    }

    public IExcelWorkbook CreateFrom(string templatePath)
    {
        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found.", templatePath);

        try
        {
            var workBook = _application.Workbooks.Add(System.IO.Path.GetFullPath(templatePath));
            var book = new ExcelWorkbook(workBook);
            return book;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create workbook from template.", ex);
        }
    }

    public IExcelWorkbook Open(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("Excel file not found.", filePath);

        try
        {
            var workBook = _application.Workbooks.Open(System.IO.Path.GetFullPath(filePath));
            var book = new ExcelWorkbook(workBook);

            return book;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to open workbook.", ex);
        }
    }

    /// <summary>
    /// 打开工作簿
    /// </summary>
    /// <param name="filename">文件路径</param>
    /// <param name="updateLinks">是否更新链接</param>
    /// <param name="readOnly">是否只读</param>
    /// <param name="format">文件格式</param>
    /// <param name="password">打开密码</param>
    /// <param name="writeResPassword">写入密码</param>
    /// <param name="ignoreReadOnlyRecommended">是否忽略只读建议</param>
    /// <param name="origin">文本来源</param>
    /// <param name="delimiter">文本分隔符</param>
    /// <param name="editable">是否可编辑</param>
    /// <param name="notify">是否通知</param>
    /// <param name="converter">格式转换器</param>
    /// <param name="addToMru">是否添加到最近使用文件</param>
    /// <returns>打开的工作簿对象</returns>
    public IExcelWorkbook OpenWorkbook(string filename, int updateLinks = 0, bool readOnly = false,
                                     int format = 1, string password = "", string writeResPassword = "",
                                     bool ignoreReadOnlyRecommended = false, int origin = 0,
                                     string delimiter = ",", bool editable = true, bool notify = false,
                                     int converter = 0, bool addToMru = true)
    {
        if (_application?.Workbooks == null || string.IsNullOrEmpty(filename))
            return null;

        try
        {
            var workbook = _application.Workbooks.Open(
                filename, updateLinks, readOnly, format, password, writeResPassword,
                ignoreReadOnlyRecommended, origin, delimiter, editable, notify,
                converter, addToMru, Type.Missing, Type.Missing) as MsExcel.Workbook;

            _performanceStats.FileOperations++;
            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 新建工作簿
    /// </summary>
    /// <param name="template">模板文件路径</param>
    /// <returns>新建的工作簿对象</returns>
    public IExcelWorkbook NewWorkbook(string template = "")
    {
        if (_application?.Workbooks == null)
            return null;

        try
        {
            MsExcel.Workbook workbook;
            if (string.IsNullOrEmpty(template))
            {
                workbook = _application.Workbooks.Add() as MsExcel.Workbook;
            }
            else
            {
                workbook = _application.Workbooks.Add(template) as MsExcel.Workbook;
            }

            _performanceStats.FileOperations++;
            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 保存所有工作簿
    /// </summary>
    public void SaveAll()
    {
        _application?.ActiveWorkbook?.Save();
        _performanceStats.FileOperations++;
    }

    /// <summary>
    /// 关闭所有工作簿
    /// </summary>
    /// <param name="saveChanges">是否保存更改</param>
    public void CloseAllWorkbooks(bool saveChanges = true)
    {
        _application?.Workbooks?.Close();
        _performanceStats.FileOperations++;
    }

    #endregion

    #region 宏和自动化
    /// <summary>
    /// 在 Excel 的撤销列表中注册一个自定义的撤销操作。
    /// 当用户选择此操作时，Excel 将尝试运行指定的宏。
    /// </summary>
    /// <param name="undoText">显示在 Excel 撤销列表中的文本。</param>
    /// <param name="macroProcedureName">
    /// 当用户选择撤销时要执行的宏的完整名称。
    /// 通常格式为 "WorkbookName.xlam!MacroName"。
    /// </param>
    /// <exception cref="System.ArgumentNullException">
    /// 如果 undoText 或 macroProcedureName 为 null 或空。
    /// </exception>
    /// <exception cref="System.InvalidOperationException">
    /// 如果内部的 Application 对象为 null。
    /// </exception>
    /// <exception cref="System.Runtime.InteropServices.COMException">
    /// 如果与 Excel 的交互失败（例如，宏名称格式错误），可能会抛出 COM 异常。
    /// </exception>
    /// <remarks>
    /// 调用此方法后，必须确保指定的 macroProcedureName 对应的宏存在于 Excel 会话中，
    /// 并且该宏实现了相应的撤销逻辑。
    /// </remarks>
    public void OnUndo(string undoText, string macroProcedureName)
    {
        if (_application == null)
        {
            throw new InvalidOperationException("Cannot register undo action: underlying Application object is null.");
        }

        // 检查参数
        if (string.IsNullOrWhiteSpace(undoText))
        {
            throw new ArgumentNullException(nameof(undoText), "Undo text cannot be null or whitespace.");
        }
        if (string.IsNullOrWhiteSpace(macroProcedureName))
        {
            throw new ArgumentNullException(nameof(macroProcedureName), "Macro procedure name cannot be null or whitespace.");
        }

        try
        {
            _application.OnUndo(undoText, macroProcedureName);
        }
        catch (COMException comEx)
        {
            log.Error($"COM Exception in RegisterUndoAction: {comEx.Message}", comEx);
            throw;
        }
        catch (Exception ex)
        {
            log.Error($"General Exception in RegisterUndoAction: {ex.Message}", ex);
            throw new InvalidOperationException("Failed to register undo action.", ex);
        }
    }

    /// <summary>
    /// 运行宏
    /// </summary>
    /// <param name="macroName">宏名称</param>
    /// <param name="args">宏参数</param>
    /// <returns>宏执行结果</returns>
    public object Run(string macroName, params object[] args)
    {
        if (_application == null || string.IsNullOrEmpty(macroName))
            return null;

        try
        {
            object result = _application.Run(macroName, args);
            _performanceStats.MacroExecutions++;
            return result;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 执行Excel 4.0宏函数
    /// </summary>
    /// <param name="macro">宏函数</param>
    /// <returns>执行结果</returns>
    public object ExecuteExcel4Macro(string macro)
    {
        if (_application == null || string.IsNullOrEmpty(macro))
            return null;

        try
        {
            object result = _application.ExecuteExcel4Macro(macro);
            _performanceStats.MacroExecutions++;
            return result;
        }
        catch
        {
            return null;
        }
    }
    #endregion

    #region 系统信息

    /// <summary>
    /// 获取操作系统版本
    /// </summary>
    public string OperatingSystem => _application?.OperatingSystem?.ToString();

    /// <summary>
    /// 获取系统内存信息
    /// </summary>
    public int MemoryFree => _application?.MemoryFree ?? 0;

    /// <summary>
    /// 获取总内存信息
    /// </summary>
    public int MemoryTotal => _application?.MemoryTotal ?? 0;

    /// <summary>
    /// 获取启动时间
    /// </summary>
    public DateTime StartupTime => _startupTime;

    /// <summary>
    /// 获取运行时间（秒）
    /// </summary>
    public double RunTime => (DateTime.Now - _startupTime).TotalSeconds;

    #endregion

    #region 国际化设置   

    /// <summary>
    /// 获取与应用程序国际设置相关的信息。
    /// </summary>
    /// <param name="index">指定要获取的国际设置项。</param>
    /// <returns>与指定索引对应的国际设置值。</returns>
    public string? International(XlApplicationInternational index)
    {
        if (_application == null)
        {
            log.Error("Underlying Application object is null in International indexer getter.");
            throw new ArgumentNullException(nameof(_application), "Cannot get International property on a null Application object.");
        }

        try
        {
            return _application.International[index]?.ToString();
        }
        catch (COMException comEx)
        {
            log.Error($"COM Exception in International indexer getter [index={index}]", comEx);
            throw;
        }
        catch (Exception ex)
        {
            log.Error($"General Exception in International indexer getter [index={index}]", ex);
            throw new InvalidOperationException($"Failed to get International property for index {index}.", ex);
        }
    }

    /// <summary>
    /// 获取或设置测量单位
    /// </summary>
    public int MeasurementUnit
    {
        get => _application != null ? _application.MeasurementUnit : 0;
        set
        {
            if (_application != null)
                _application.MeasurementUnit = value;
        }
    }

    /// <summary>
    /// 获取或设置默认文件路径
    /// </summary>
    public string DefaultFilePath
    {
        get => _application?.DefaultFilePath?.ToString();
        set
        {
            if (_application != null && value != null)
                _application.DefaultFilePath = value;
        }
    }

    /// <summary>
    /// 获取或设置模板路径
    /// </summary>
    public string TemplatesPath => _application?.TemplatesPath?.ToString();

    #endregion

    #region 剪贴板操作

    /// <summary>
    /// 获取或设置剪切复制模式
    /// </summary>
    public XlCutCopyMode CutCopyMode
    {
        get => _application != null ? (XlCutCopyMode)_application.CutCopyMode : XlCutCopyMode.xlCopy;
        set
        {
            if (_application != null)
                _application.CutCopyMode = (MsExcel.XlCutCopyMode)value;
        }
    }

    /// <summary>
    /// 清除剪贴板
    /// </summary>
    public void ClearClipboard()
    {
        if (_application != null)
            _application.CutCopyMode = MsExcel.XlCutCopyMode.xlCopy;
    }

    /// <summary>
    /// 获取剪贴板内容类型
    /// </summary>
    public string ClipboardContentType => "Unknown"; // Excel Application不直接暴露此属性

    #endregion

    #region 对话框和用户界面

    /// <summary>
    /// 显示"打开"对话框获取文件名
    /// </summary>
    /// <param name="filter">文件过滤器（如："Excel Files (*.xlsx), *.xlsx|All Files (*.*), *.*"）</param>
    /// <param name="filterIndex">默认过滤器索引</param>
    /// <param name="title">对话框标题</param>
    /// <param name="buttonText">按钮文本（仅Mac）</param>
    /// <param name="multiSelect">是否允许多选</param>
    /// <returns>选择的文件路径（多选时为数组），取消时为null</returns>
    public IList<string> GetOpenFilenames(
        string filter = "All Files (*.*), *.*",
        int filterIndex = 1,
        string title = "Open",
        string buttonText = "Open",
        bool multiSelect = true)
    {
        try
        {
            object result = _application.GetOpenFilename(
                FileFilter: filter,
                FilterIndex: filterIndex,
                Title: title,
                ButtonText: buttonText,
                MultiSelect: multiSelect
            );

            return ParseFilenameResult(result, multiSelect);
        }
        catch (COMException ex)
        {
            throw new ExcelOperationException("Error getting open filename", ex);
        }
    }

    private IList<string> ParseFilenameResult(object result, bool multiSelect)
    {
        if (result == null || result.Equals(false))
            return null;

        if (!multiSelect)
        {
            return [result.ToString()];
        }

        if (result is object[,] multiResult)
        {
            var files = new List<string>();
            int count = multiResult.GetLength(0);

            for (int i = 1; i <= count; i++) // Excel 数组是基1索引
            {
                files.Add(multiResult[i, 1]?.ToString());
            }

            return files;
        }

        throw new InvalidCastException("Unexpected return type from GetOpenFilename");
    }

    /// <summary>
    /// 显示文件打开对话框
    /// </summary>
    /// <param name="fileFilter">文件过滤器</param>
    /// <param name="title">对话框标题</param>
    /// <returns>选择的文件路径</returns>
    public object GetOpenFilename(string fileFilter = "所有文件 (*.*)|*.*",
                                   string title = "打开文件")
    {
        if (_application == null) return null;

        try
        {
            return _application.GetOpenFilename(
                FileFilter: fileFilter,
                MultiSelect: false,
                Title: title);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 显示"另存为"对话框获取文件名
    /// </summary>
    /// <param name="initialFilename">初始文件名</param>
    /// <param name="filter">文件过滤器</param>
    /// <param name="filterIndex">默认过滤器索引</param>
    /// <param name="title">对话框标题</param>
    /// <param name="buttonText">按钮文本（仅Mac）</param>
    /// <returns>选择的文件路径，取消时为null</returns>
    public string GetSaveAsFilename(
        string initialFilename = "",
        string filter = "Excel文件 (*.xlsx)|*.xlsx|Excel 97-2003文件 (*.xls)|*.xls|所有文件 (*.*)|*.*",
        int filterIndex = 1,
        string title = "另存为",
        string buttonText = "保存")
    {
        try
        {
            object result = _application.GetSaveAsFilename(
                InitialFilename: initialFilename,
                FileFilter: filter,
                FilterIndex: filterIndex,
                Title: title,
                ButtonText: buttonText
            );

            return result as string;
        }
        catch (COMException ex)
        {
            throw new ExcelOperationException("Error getting save as filename", ex);
        }
    }

    /// <summary>
    /// 显示文件保存对话框
    /// </summary>
    /// <param name="initialFilename">初始文件名</param>
    /// <param name="fileFilter">文件过滤器</param>
    /// <param name="title">对话框标题</param>
    /// <returns>保存的文件路径</returns>
    public string GetSaveFilename(string initialFilename = "",
                                   string fileFilter = "Excel文件 (*.xlsx)|*.xlsx|Excel 97-2003文件 (*.xls)|*.xls|所有文件 (*.*)|*.*",
                                   string title = "另存为")
    {
        if (_application == null) return null;

        try
        {
            object result = _application.GetSaveAsFilename(
                                        InitialFilename: initialFilename,
                                        FileFilter: fileFilter,
                                        FilterIndex: 1,
                                        ButtonText: "保存",
                                        Title: title);
            return result?.ToString();
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 获取自定义序列的序号
    /// </summary>
    /// <param name="listItems">序列项目</param>
    /// <returns>序列序号（未找到返回-1）</returns>
    public int GetCustomListNum(IList<string> listItems)
    {
        if (listItems == null || listItems.Count == 0)
            throw new ArgumentException("List items cannot be null or empty", nameof(listItems));

        try
        {
            // 将列表转换为对象数组（需要基1索引）
            object[,] array = new object[listItems.Count, 1];
            for (int i = 0; i < listItems.Count; i++)
            {
                array[i, 0] = listItems[i];
            }

            return _application.GetCustomListNum(array);
        }
        catch (COMException ex)
        {
            // 未找到自定义列表时返回特定错误
            if (ex.ErrorCode == -2146827284) // 0x800A03EC
                return -1;

            throw new ExcelOperationException("Error getting custom list number", ex);
        }
    }

    /// <summary>
    /// 获取自定义序列的内容
    /// </summary>
    /// <param name="listNum">序列序号</param>
    /// <returns>序列内容数组</returns>
    public IList<string> GetCustomListContents(int listNum)
    {
        if (listNum < 1 || listNum > _application.CustomListCount)
            throw new ArgumentOutOfRangeException(nameof(listNum), "List number is out of range");

        try
        {
            object result = _application.GetCustomListContents(listNum);

            // 处理返回结果（可能是一维或二维数组）
            if (result is object[,] array2D)
            {
                var items = new List<string>();
                int rows = array2D.GetLength(0);

                for (int i = 1; i <= rows; i++) // Excel 数组是基1索引
                {
                    items.Add(array2D[i, 1]?.ToString());
                }

                return items;
            }

            if (result is string[] array1D)
            {
                return array1D;
            }

            throw new InvalidCastException("Unexpected return type from GetCustomListContents");
        }
        catch (COMException ex)
        {
            throw new ExcelOperationException("Error getting custom list contents", ex);
        }
    }


    /// <summary>
    /// 显示一个提示用户输入信息的对话框。
    /// 这是对 Microsoft.Office.Interop.Excel.Application.InputBox 方法的封装。
    /// </summary>
    /// <param name="prompt">显示在对话框中的消息 (最多 255 个字符)。</param>
    /// <param name="title">对话框标题栏的文本。如果省略，则使用应用程序名称。</param>
    /// <param name="defaultValue">文本框中的默认值。如果省略，则文本框为空。</param>
    /// <param name="left">对话框距屏幕左边的距离（以点为单位）。如果省略，则 Excel 设置对话框位置。</param>
    /// <param name="top">对话框距屏幕上边的距离（以点为单位）。如果省略，则 Excel 设置对话框位置。</param>
    /// <param name="helpFile">帮助文件的名称。如果省略，则不显示帮助。</param>
    /// <param name="helpContextID">帮助文件中帮助主题的上下文编号。如果省略，则不显示帮助。</param>
    /// <param name="type">
    /// 指定在对话框中返回的数据类型。默认值为 2 (文本)。
    /// 可以是以下值的和：
    /// 0 - 公式, 1 - 数字, 2 - 文本, 4 - 逻辑值, 8 - 单元格引用, 16 - 错误值, 64 - 数组。
    /// 如果类型为 8 (单元格引用)，则返回一个 Range 对象。
    /// </param>
    /// <returns>
    /// 一个 InputBoxResult 对象，指示用户是点击了确定还是取消，以及返回的值。
    /// 如果用户点击取消，ResultType 为 Cancel，Value 为 null。
    /// 如果发生错误，ResultType 为 Error，Value 可能包含错误信息。
    /// 如果用户点击确定，ResultType 为 Ok，Value 包含用户输入的值（类型取决于 type 参数）。
    /// </returns>
    /// <exception cref="ArgumentNullException">
    /// 如果内部的 _application 对象为 null。
    /// </exception>
    public InputBoxResult ShowInputBox(
        string prompt,
        string? title = null,
        object? defaultValue = null,
        object? left = null,
        object? top = null,
        object? helpFile = null,
        object? helpContextID = null,
        object? type = null)
    {
        if (_application == null)
        {
            // 可以根据你的错误处理策略抛出异常或返回特定值
            log.Error("Underlying Application object is null.");
            // throw new ArgumentNullException(nameof(_application), "Underlying Application object is null.");
            // 或者返回错误结果
            return new InputBoxResult(InputBoxResultType.Error, "Application object is null.");
        }

        // 确保 Prompt 不为 null 或空
        if (string.IsNullOrEmpty(prompt))
        {
            log.Warn("InputBox prompt is null or empty.");
            // 可以抛出异常或使用默认提示
            prompt = "请输入内容:"; // 使用默认提示
        }

        try
        {
            // 使用 Type.Missing 表示省略的可选参数，除非调用者提供了具体值
            object titleParam = title ?? Type.Missing;
            object defaultParam = defaultValue ?? Type.Missing;
            object leftParam = left ?? Type.Missing;
            object topParam = top ?? Type.Missing;
            object helpFileParam = helpFile ?? Type.Missing;
            object helpContextIDParam = helpContextID ?? Type.Missing;
            object typeParam = type ?? Type.Missing; // 默认是 2 (文本)

            // 调用 Interop 的 InputBox 方法
            object result = _application.InputBox(
                prompt,
                titleParam,
                defaultParam,
                leftParam,
                topParam,
                helpFileParam,
                helpContextIDParam,
                typeParam
            );

            // 检查返回值是否是用户取消操作的标志
            // 注意：InputBox 在用户点击取消时通常会抛出异常，而不是返回特定值。
            // 因此，这个检查可能不是必需的，除非 Interop 包装有特殊处理。
            // 但我们保留它作为一种健壮性检查。
            // if (result is double doubleResult && doubleResult == 0.0 && typeParam is double typeDouble && typeDouble == 8)
            // {
            //     // 特殊情况：Type=8 且返回 0.0 可能表示取消？需要验证
            //     // 通常还是依赖异常捕获
            // }

            // 如果没有异常抛出，通常表示用户点击了确定
            return new InputBoxResult(InputBoxResultType.Ok, result);

        }
        // 捕获用户点击取消按钮时可能抛出的特定 COM 异常
        catch (COMException comEx) when (comEx.HResult == unchecked((int)0x800A03EC)) // DISP_E_EXCEPTION
        {
            log.Info("User cancelled the InputBox.");
            // 用户点击了取消按钮，返回 Cancel 结果
            return new InputBoxResult(InputBoxResultType.Cancel, null);
        }
        catch (COMException comEx)
        {
            // 记录或重新抛出其他 COM 异常
            log.Error($"COM Exception in ShowInputBox: {comEx.Message}", comEx);
            // 可以选择重新抛出或返回错误结果
            return new InputBoxResult(InputBoxResultType.Error, $"COM Error: {comEx.Message}");
        }
        catch (Exception ex) // 捕获其他可能的异常
        {
            log.Error($"General Exception in ShowInputBox: {ex.Message}", ex);
            // 根据需要决定是抛出还是处理
            // throw new InvalidOperationException("Failed to show InputBox.", ex);
            return new InputBoxResult(InputBoxResultType.Error, $"Error: {ex.Message}");
        }
    }


    /// <summary>
    /// 显示一个提示用户输入文本的对话框。
    /// </summary>
    /// <param name="prompt">显示在对话框中的消息。</param>
    /// <param name="title">对话框标题栏的文本。</param>
    /// <param name="defaultValue">文本框中的默认文本。</param>
    /// <returns>包含用户输入文本的 InputBoxResult。</returns>
    public InputBoxResult<string?> ShowInputBoxText(string prompt, string? title = null, string? defaultValue = null)
    {
        object defaultVal = defaultValue ?? Type.Missing;
        var result = ShowInputBox(prompt, title, type: 2);
        return new InputBoxResult<string?>(result.ResultType, result.Value?.ToString());
    }

    /// <summary>
    /// 显示一个提示用户输入数字的对话框。
    /// </summary>
    /// <param name="prompt">显示在对话框中的消息。</param>
    /// <param name="title">对话框标题栏的文本。</param>
    /// <param name="defaultValue">文本框中的默认数字。</param>
    /// <returns>包含用户输入数字的 InputBoxResult。</returns>
    public InputBoxResult<double?> ShowInputBoxNumber(string prompt, string? title = null, double? defaultValue = null)
    {
        object defaultVal = defaultValue ?? Type.Missing;
        var result = ShowInputBox(prompt, title, type: 1);
        if (result.Value != null && result.ResultType == InputBoxResultType.Ok)
        {
            var dVal = Convert.ToDouble(result.Value);
            return new InputBoxResult<double?>(result.ResultType, dVal);
        }
        return new InputBoxResult<double?>(InputBoxResultType.Error, null);
    }

    /// <summary>
    /// 显示一个提示用户选择单元格引用的对话框。
    /// </summary>
    /// <param name="prompt">显示在对话框中的消息。</param>
    /// <param name="title">对话框标题栏的文本。</param>
    /// <returns>包含用户选择的 Range 对象的 InputBoxResult。</returns>
    public InputBoxResult<IExcelRange> ShowInputBoxRangeSelection(string prompt, string? title = null)
    {
        var result = ShowInputBox(prompt, title, type: 8);
        if (result.Value is MsExcel.Range rang && rang != null)
        {
            var obj = new ExcelRange(rang);
            return new InputBoxResult<IExcelRange>(result.ResultType, obj);
        }
        return new InputBoxResult<IExcelRange>(InputBoxResultType.Error, null);
    }

    /// <summary>
    /// 显示一个提示用户输入信息的对话框，并返回 object 类型的结果。
    /// 注意：如果用户点击取消，通常会抛出 COMException (HResult = 0x800A03EC)。
    /// </summary>
    /// <param name="prompt">显示在对话框中的消息。</param>
    /// <param name="title">对话框标题栏的文本。</param>
    /// <param name="defaultValue">文本框中的默认值。</param>
    /// <param name="left">对话框距屏幕左边的距离。</param>
    /// <param name="top">对话框距屏幕上边的距离。</param>
    /// <param name="helpFile">帮助文件的名称。</param>
    /// <param name="helpContextID">帮助文件中帮助主题的上下文编号。</param>
    /// <param name="type">
    /// 指定在对话框中返回的数据类型。默认值为 2 (文本)。
    /// 可以是以下值的和：
    /// 0 - 公式, 1 - 数字, 2 - 文本, 4 - 逻辑值, 8 - 单元格引用, 16 - 错误值, 64 - 数组。
    /// 如果类型为 8 (单元格引用)，则返回一个 Range 对象。
    /// </param>
    /// <returns>用户输入的值。类型取决于 type 参数。如果用户取消，则抛出 COMException。</returns>
    public object InputBox(
        string prompt,
        object? title = null,
        object? defaultValue = null,
        object? left = null,
        object? top = null,
        object? helpFile = null,
        object? helpContextID = null,
        int? type = 2)
    {
        if (_application == null)
            throw new ArgumentNullException(nameof(_application), "Underlying Application object is null.");

        if (string.IsNullOrEmpty(prompt))
            prompt = "请输入内容:";

        var obj = _application.InputBox(
                    prompt,
                    title ?? Type.Missing,
                    defaultValue ?? Type.Missing,
                    left ?? Type.Missing,
                    top ?? Type.Missing,
                    helpFile ?? Type.Missing,
                    helpContextID ?? Type.Missing,
                    type ?? Type.Missing
        );

        if (type == 8 && obj is MsExcel.Range rang)
            return new ExcelRange(rang);
        return obj;
    }
    #endregion

    #region 打印设置

    /// <summary>
    /// 获取或设置是否使用系统分隔符
    /// </summary>
    public bool UseSystemSeparators
    {
        get => _application != null && _application.UseSystemSeparators;
        set
        {
            if (_application != null)
                _application.UseSystemSeparators = value;
        }
    }

    /// <summary>
    /// 获取或设置默认打印机
    /// </summary>
    public string DefaultPrinter
    {
        get => _application?.ActivePrinter?.ToString();
        set
        {
            if (_application != null && value != null)
                _application.ActivePrinter = value;
        }
    }

    /// <summary>
    /// 获取打印机列表
    /// </summary>
    /// <returns>打印机名称数组</returns>
    public string[] GetPrinterList()
    {
        // Excel Application不直接提供打印机列表
        return new string[0];
    }

    /// <summary>
    /// 打印预览所有工作簿
    /// </summary>
    public void PrintPreviewAll()
    {
        _application?.ActiveWorkbook?.PrintPreview();
    }

    #endregion

    #region 错误处理

    /// <summary>
    /// 获取或设置错误检查选项
    /// </summary>
    public IExcelErrorCheckingOptions ErrorCheckingOptions => _errorCheckingOptions ??= new ExcelErrorCheckingOptions(_application?.ErrorCheckingOptions);

    #endregion

    #region 性能监控

    /// <summary>
    /// 获取性能统计信息
    /// </summary>
    /// <returns>性能统计对象</returns>
    public ApplicationPerformance GetPerformanceStats()
    {
        _performanceStats.RunTime = RunTime;
        return _performanceStats;
    }

    /// <summary>
    /// 重置性能统计
    /// </summary>
    public void ResetPerformanceStats()
    {
        _performanceStats = new ApplicationPerformance { StartupTime = _startupTime };
    }

    /// <summary>
    /// 获取内存使用情况
    /// </summary>
    /// <returns>内存使用信息</returns>
    public MemoryInfo GetMemoryInfo()
    {
        var memoryInfo = new MemoryInfo
        {
            FreeMemory = MemoryFree,
            TotalMemory = MemoryTotal,
            UsedMemory = MemoryTotal - MemoryFree
        };

        memoryInfo.UsagePercentage = memoryInfo.TotalMemory > 0 ?
            (double)memoryInfo.UsedMemory / memoryInfo.TotalMemory * 100 : 0;

        return memoryInfo;
    }

    /// <summary>
    /// 获取CPU使用率
    /// </summary>
    /// <returns>CPU使用率百分比</returns>
    public double GetCPUUsage()
    {
        // Excel Application不直接提供CPU使用率
        return 0;
    }

    #endregion

    #region 操作方法
    public IExcelRange? Range(object? cell1, object? cell2 = null)
    {
        if (_application == null)
            return null;
        if (cell1 is ExcelRange range1)
            cell1 = range1.Range;
        if (cell2 is ExcelRange range2)
            cell2 = range2.Range;
        cell1 ??= Type.Missing;
        cell2 ??= Type.Missing;
        return new ExcelRange(_application.Range[cell1, cell2]);
    }



    /// <summary>
    /// 选定指定的区域或对象。
    /// </summary>
    /// <param name="reference">要选定的区域或对象（可以是 Range, Sheet 名称等）。</param>
    /// <param name="scroll">是否滚动到选定区域。</param>
    public void Goto(object reference, bool scroll = true)
    {
        if (_application == null) return;
        try
        {
            // reference 可能是 MsExcel.Range, string (range address), 或其他对象
            _application.Goto(reference, scroll);
            _performanceStats.MiscOperations++; // 假设有此计数器或添加一个
        }
        catch (Exception ex)
        {
            // 根据需要记录日志或重新抛出
            System.Diagnostics.Debug.WriteLine($"Goto failed: {ex.Message}");
            // throw; // 可选：重新抛出异常
        }
    }

    /// <summary>
    /// 将公式从一种引用样式转换为另一种。
    /// </summary>
    /// <param name="formula">要转换的公式。</param>
    /// <param name="fromReferenceStyle">源引用样式。</param>
    /// <param name="toReferenceStyle">目标引用样式。</param>
    /// <param name="toAbsolute">如何转换引用（绝对、相对等）。</param>
    /// <param name="relativeTo">相对引用的基准单元格。</param>
    /// <returns>转换后的公式。</returns>
    public string ConvertFormula(string formula, XlReferenceStyle fromReferenceStyle,
                                 XlReferenceStyle toReferenceStyle, int toAbsolute = 1, // XlAbsolute (1) is default
                                 object? relativeTo = null)
    {
        if (_application == null || string.IsNullOrEmpty(formula))
            return formula;

        try
        {
            object comRelativeTo = relativeTo;
            // 如果 relativeTo 是 ExcelRange, 需要提取内部 MsExcel.Range
            if (relativeTo is ExcelRange excelRange)
            {
                comRelativeTo = excelRange.InternalRange;
            }
            // 如果 relativeTo 是 ExcelWorksheet, 需要提取内部 MsExcel.Worksheet 的 Range? 通常 relativeTo 是一个 Range
            else if (relativeTo is ExcelWorksheet excelWorksheet)
            {
                comRelativeTo = excelWorksheet.Worksheet; // 不确定是否正确，通常需要 Range
            }

            object result = _application.ConvertFormula(formula,
                                                        (MsExcel.XlReferenceStyle)fromReferenceStyle,
                                                        (MsExcel.XlReferenceStyle)toReferenceStyle,
                                                        toAbsolute,
                                                        comRelativeTo ?? Type.Missing);
            _performanceStats.MiscOperations++;
            return result?.ToString() ?? formula;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"ConvertFormula failed: {ex.Message}");
            return formula; // 返回原始公式或抛出异常
        }
    }

    /// <summary>
    /// 检查指定文本的拼写。
    /// </summary>
    /// <param name="text">要检查的文本。</param>
    /// <param name="customDictionary">自定义词典的名称。</param>
    /// <param name="ignoreUpper">是否忽略全大写单词。</param>
    /// <returns>如果拼写正确返回 True，否则返回 False。</returns>
    public bool CheckSpelling(string text, object? customDictionary = null, object? ignoreUpper = null)
    {
        if (_application == null || string.IsNullOrEmpty(text))
            return true; // 或 false，取决于默认行为假设

        try
        {
            bool result = _application.CheckSpelling(text, customDictionary ?? Type.Missing, ignoreUpper ?? Type.Missing);
            _performanceStats.MiscOperations++;
            return result;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"CheckSpelling failed: {ex.Message}");
            return true; // 假设检查失败时返回正确，或根据需要处理
        }
    }

    /// <summary>
    /// 为指定的键或键组合指定过程（宏）。
    /// </summary>
    /// <param name="key">键或键组合（例如 "^c" 代表 Ctrl+C）。</param>
    /// <param name="procedure">要运行的过程名称（宏名）。</param>
    public void OnKey(string key, string procedure = "")
    {
        if (_application == null || string.IsNullOrEmpty(key)) return;

        try
        {
            // procedure 为空字符串时，会删除之前的 OnKey 设置
            _application.OnKey(key, procedure);
            _performanceStats.MiscOperations++;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"OnKey failed: {ex.Message}");
        }
    }
    /// <summary>
    /// 最小化应用程序
    /// </summary>
    public void Minimize()
    {
        if (_application != null)
            _application.WindowState = MsExcel.XlWindowState.xlMinimized;
    }

    /// <summary>
    /// 最大化应用程序
    /// </summary>
    public void Maximize()
    {
        if (_application != null)
            _application.WindowState = MsExcel.XlWindowState.xlMaximized;
    }

    /// <summary>
    /// 恢复应用程序
    /// </summary>
    public void Restore()
    {
        if (_application != null)
            _application.WindowState = MsExcel.XlWindowState.xlNormal;
    }

    /// <summary>
    /// 退出应用程序
    /// </summary>
    /// <param name="saveChanges">是否保存更改</param>
    public void Quit(bool saveChanges = true)
    {
        if (_application == null) return;

        try
        {
            _application.Quit();
        }
        catch
        {
            // 忽略退出过程中的异常
        }
    }

    /// <summary>
    /// 发送按键到应用程序
    /// </summary>
    /// <param name="keys">按键字符串</param>
    /// <param name="wait">是否等待</param>
    public void SendKeys(string keys, bool wait = true)
    {
        if (_application == null || string.IsNullOrEmpty(keys)) return;

        try
        {
            _application.SendKeys(keys, wait);
        }
        catch
        {
            // 忽略发送按键过程中的异常
        }
    }

    /// <summary>
    /// 等待指定时间
    /// </summary>
    /// <param name="time">等待到的时间</param>
    public void Wait(DateTime time)
    {
        if (_application == null) return;

        try
        {
            _application.Wait(time.ToOADate());
        }
        catch
        {
            // 忽略等待过程中的异常
        }
    }

    /// <summary>
    /// 延迟指定毫秒数
    /// </summary>
    /// <param name="milliseconds">毫秒数</param>
    public void Delay(int milliseconds)
    {
        if (milliseconds <= 0) return;

        try
        {
            DateTime waitTime = DateTime.Now.AddMilliseconds(milliseconds);
            Wait(waitTime);
        }
        catch
        {
            // 忽略延迟过程中的异常
        }
    }
    #endregion

    #region 事件实现

    public event WindowResizeEventHandler WindowResize
    {
        add { _windowResize += value; }
        remove { _windowResize -= value; }
    }

    public event WindowDeactivateEventHandler WindowDeactivate
    {
        add { _windowDeactivate += value; }
        remove { _windowDeactivate -= value; }
    }

    public event WindowActivateEventHandler WindowActivate
    {
        add { _windowActivate += value; }
        remove { _windowActivate -= value; }
    }

    public event WorkbookNewEventHandler WorkbookNew
    {
        add { _newWorkbook += value; }
        remove { _newWorkbook -= value; }
    }

    /// <summary>
    /// 当工作簿打开时触发
    /// </summary>
    public event WorkbookOpenEventHandler WorkbookOpen
    {
        add { _workbookOpen += value; }
        remove { _workbookOpen -= value; }
    }

    /// <summary>
    /// 当工作簿被激活时触发
    /// </summary>
    public event WorkbookActivateEventHandler WorkbookActivate
    {
        add { _workbookActivate += value; }
        remove { _workbookActivate -= value; }
    }

    /// <summary>
    /// 当工作簿被取消激活时触发
    /// </summary>
    public event WorkbookDeactivateEventHandler WorkbookDeactivate
    {
        add { _workbookDeactivate += value; }
        remove { _workbookDeactivate -= value; }
    }

    /// <summary>
    /// 当工作簿即将关闭时触发
    /// </summary>
    public event WorkbookBeforeCloseEventHandler WorkbookBeforeClose
    {
        add { _workbookBeforeClose += value; }
        remove { _workbookBeforeClose -= value; }
    }

    public event WorkbookBeforeSaveEventHandler WorkbookBeforeSave
    {
        add { _workbookBeforeSave += value; }
        remove { _workbookBeforeSave -= value; }
    }

    /// <summary>
    /// 当工作表内容发生改变时触发
    /// </summary>
    public event SheetChangeEventHandler SheetChange
    {
        add { _sheetChange += value; }
        remove { _sheetChange -= value; }
    }

    /// <summary>
    /// 当工作表被激活时触发
    /// </summary>
    public event SheetActivateEventHandler SheetActivate
    {
        add { _sheetActivate += value; }
        remove { _sheetActivate -= value; }
    }

    /// <summary>
    /// 当工作表被取消激活时触发
    /// </summary>
    public event SheetDeactivateEventHandler SheetDeactivate
    {
        add { _sheetDeactivate += value; }
        remove { _sheetDeactivate -= value; }
    }

    /// <summary>
    /// 当工作表选择区域发生改变时触发
    /// </summary>
    public event SheetSelectionChangeEventHandler SheetSelectionChange
    {
        add { _sheetSelectionChange += value; }
        remove { _sheetSelectionChange -= value; }
    }

    /// <summary>
    /// 当工作表被双击前触发
    /// </summary>
    public event SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClick
    {
        add { _sheetBeforeDoubleClick += value; }
        remove { _sheetBeforeDoubleClick -= value; }
    }

    /// <summary>
    /// 当工作表被右键单击前触发
    /// </summary>
    public event SheetBeforeRightClickEventHandler SheetBeforeRightClick
    {
        add { _sheetBeforeRightClick += value; }
        remove { _sheetBeforeRightClick -= value; }
    }

    /// <summary>
    /// 当工作表计算完成时触发
    /// </summary>
    public event SheetCalculateEventHandler SheetCalculate
    {
        add { _sheetCalculate += value; }
        remove { _sheetCalculate -= value; }
    }

    #endregion

    #region 事件处理方法

    private void OnWindowResize(MsExcel.Workbook Wb, MsExcel.Window Wn)
    {
        if (_windowResize != null && Wb != null)
        {
            var excelWorkbook = new ExcelWorkbook(Wb);
            var excelwindows = Wn != null ? new ExcelWindow(Wn, excelWorkbook) : null;
            _windowResize(excelWorkbook, excelwindows);
        }
    }

    private void OnWindowDeactivate(MsExcel.Workbook Wb, MsExcel.Window Wn)
    {
        if (_windowDeactivate != null && Wb != null)
        {
            var excelWorkbook = new ExcelWorkbook(Wb);
            var excelwindows = Wn != null ? new ExcelWindow(Wn, excelWorkbook) : null;
            _windowDeactivate(excelWorkbook, excelwindows);
        }
    }

    private void OnWindowActivate(MsExcel.Workbook Wb, MsExcel.Window Wn)
    {
        if (_windowActivate != null && Wb != null)
        {
            var excelWorkbook = new ExcelWorkbook(Wb);
            var excelwindows = Wn != null ? new ExcelWindow(Wn, excelWorkbook) : null;
            _windowActivate(excelWorkbook, excelwindows);
        }
    }

    private void OnNewWorkbook(MsExcel.Workbook workbook)
    {
        if (_newWorkbook != null && workbook != null)
        {
            var excelWorkbook = new ExcelWorkbook(workbook);
            _newWorkbook(excelWorkbook);
        }
    }

    /// <summary>
    /// 处理WorkbookOpen事件
    /// </summary>
    /// <param name="workbook">打开的工作簿</param>
    private void OnWorkbookOpen(MsExcel.Workbook workbook)
    {
        if (_workbookOpen != null && workbook != null)
        {
            var excelWorkbook = new ExcelWorkbook(workbook);
            _workbookOpen(excelWorkbook);
        }
    }

    /// <summary>
    /// 处理WorkbookActivate事件
    /// </summary>
    /// <param name="workbook">激活的工作簿</param>
    private void OnWorkbookActivate(MsExcel.Workbook workbook)
    {
        if (_workbookActivate != null && workbook != null)
        {
            var excelWorkbook = new ExcelWorkbook(workbook);
            _workbookActivate(excelWorkbook);
        }
    }

    /// <summary>
    /// 处理WorkbookDeactivate事件
    /// </summary>
    /// <param name="workbook">取消激活的工作簿</param>
    private void OnWorkbookDeactivate(MsExcel.Workbook workbook)
    {
        if (_workbookDeactivate != null && workbook != null)
        {
            var excelWorkbook = new ExcelWorkbook(workbook);
            _workbookDeactivate(excelWorkbook);
        }
    }

    private void OnWorkbookBeforeSave(MsExcel.Workbook workbook, bool SaveAsUI, ref bool Cancel)
    {
        if (_workbookBeforeSave != null && workbook != null)
        {
            var excelWorkbook = new ExcelWorkbook(workbook);
            _workbookBeforeSave(excelWorkbook, SaveAsUI, ref Cancel);
        }
    }

    /// <summary>
    /// 处理WorkbookBeforeClose事件
    /// </summary>
    /// <param name="workbook">即将关闭的工作簿</param>
    /// <param name="cancel">是否取消关闭操作</param>
    private void OnWorkbookBeforeClose(MsExcel.Workbook workbook, ref bool cancel)
    {
        if (_workbookBeforeClose != null && workbook != null)
        {
            var excelWorkbook = new ExcelWorkbook(workbook);
            _workbookBeforeClose(excelWorkbook, ref cancel);
        }
    }

    /// <summary>
    /// 处理SheetChange事件
    /// </summary>
    /// <param name="sheet">发生变化的工作表</param>
    /// <param name="range">发生变化的单元格区域</param>
    private void OnSheetChange(object sheet, MsExcel.Range range)
    {
        _activeSheet?.Dispose();
        _activeSheet = null;
        if (_sheetChange != null && sheet != null && range != null)
        {
            var excelWorksheet = new ExcelWorksheet(sheet as MsExcel.Worksheet);
            var excelRange = new ExcelRange(range);
            _sheetChange(excelWorksheet, excelRange);
        }
    }

    /// <summary>
    /// 处理SheetActivate事件
    /// </summary>
    /// <param name="sheet">激活的工作表</param>
    private void OnSheetActivate(object sheet)
    {
        _activeSheet?.Dispose();
        _activeSheet = null;
        if (_sheetActivate != null && sheet != null)
        {
            var excelWorksheet = new ExcelWorksheet(sheet as MsExcel.Worksheet);
            _sheetActivate(excelWorksheet);
        }
    }

    /// <summary>
    /// 处理SheetDeactivate事件
    /// </summary>
    /// <param name="sheet">取消激活的工作表</param>
    private void OnSheetDeactivate(object sheet)
    {
        _activeSheet?.Dispose();
        _activeSheet = null;
        if (_sheetDeactivate != null && sheet != null)
        {
            var excelWorksheet = new ExcelWorksheet(sheet as MsExcel.Worksheet);
            _sheetDeactivate(excelWorksheet);
        }
    }

    /// <summary>
    /// 处理SheetSelectionChange事件
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="target">选中的区域</param>
    private void OnSheetSelectionChange(object sheet, MsExcel.Range target)
    {
        if (_sheetSelectionChange != null && sheet != null && target != null)
        {
            var excelWorksheet = new ExcelWorksheet(sheet as MsExcel.Worksheet);
            var excelRange = new ExcelRange(target);
            _sheetSelectionChange(excelWorksheet, excelRange);
        }
    }

    /// <summary>
    /// 处理SheetBeforeDoubleClick事件
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="target">双击的单元格</param>
    /// <param name="cancel">是否取消默认操作</param>
    private void OnSheetBeforeDoubleClick(object sheet, MsExcel.Range target, ref bool cancel)
    {
        if (_sheetBeforeDoubleClick != null && sheet != null && target != null)
        {
            var excelWorksheet = new ExcelWorksheet(sheet as MsExcel.Worksheet);
            var excelRange = new ExcelRange(target);
            _sheetBeforeDoubleClick(excelWorksheet, excelRange, ref cancel);
        }
    }

    /// <summary>
    /// 处理SheetBeforeRightClick事件
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="target">右击的单元格</param>
    /// <param name="cancel">是否取消默认操作</param>
    private void OnSheetBeforeRightClick(object sheet, MsExcel.Range target, ref bool cancel)
    {
        if (_sheetBeforeRightClick != null && sheet != null && target != null)
        {
            var excelWorksheet = new ExcelWorksheet(sheet as MsExcel.Worksheet);
            var excelRange = new ExcelRange(target);
            _sheetBeforeRightClick(excelWorksheet, excelRange, ref cancel);
        }
    }

    /// <summary>
    /// 处理SheetCalculate事件
    /// </summary>
    /// <param name="sheet">工作表</param>
    private void OnSheetCalculate(object sheet)
    {
        if (_sheetCalculate != null && sheet != null)
        {
            var excelWorksheet = new ExcelWorksheet(sheet as MsExcel.Worksheet);
            _sheetCalculate(excelWorksheet);
        }
    }

    #endregion
}