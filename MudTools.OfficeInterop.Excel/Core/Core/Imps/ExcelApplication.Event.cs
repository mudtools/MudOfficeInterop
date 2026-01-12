//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

partial class ExcelApplication
{
    private MsExcel.AppEvents_Event _appEvents_Event;

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

    /// <summary>
    /// 初始化 ExcelApplication 实例（创建新的Excel应用程序）
    /// </summary>
    public ExcelApplication()
    {
        var app = new MsExcel.Application();
        _application = app;
        _appEvents_Event = app;
        _disposedValue = false;
        InitializeEvents();
    }
    /// <summary>
    /// 初始化 ExcelApplication 实例
    /// </summary>
    /// <param name="application">底层的 COM Application 对象</param>
    internal ExcelApplication(MsExcel.Application application)
    {
        _application = application ?? throw new ArgumentNullException(nameof(application));
        _appEvents_Event = application;
        InitializeEvents();
    }

    internal ExcelApplication(MsExcel._Application application)
    {
        _application = application ?? throw new ArgumentNullException(nameof(application));
    }

    public IExcelWorkbook BlankWorkbook()
    {
        if (_application == null)
            throw new ObjectDisposedException(nameof(_application));

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
        if (_application == null)
            throw new ObjectDisposedException(nameof(_application));

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
        if (_application == null)
            throw new ObjectDisposedException(nameof(_application));
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
    public IExcelWorkbook? OpenWorkbook(string filename, int updateLinks = 0, bool readOnly = false,
                                     int format = 1, string password = "", string writeResPassword = "",
                                     bool ignoreReadOnlyRecommended = false, int origin = 0,
                                     string delimiter = ",", bool editable = true, bool notify = false,
                                     int converter = 0, bool addToMru = true)
    {
        if (_application == null)
            throw new ObjectDisposedException(nameof(_application));
        if (_application?.Workbooks == null || string.IsNullOrEmpty(filename))
            return null;

        try
        {
            var workbook = _application.Workbooks.Open(
                filename, updateLinks, readOnly, format, password, writeResPassword,
                ignoreReadOnlyRecommended, origin, delimiter, editable, notify,
                converter, addToMru, Type.Missing, Type.Missing);

            return workbook != null ? new ExcelWorkbook(workbook) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 创建文件对话框
    /// </summary>
    /// <param name="fileDialogType">文件对话框类型</param>
    /// <returns>文件对话框对象</returns>
    public IOfficeFileDialog? CreateFileDialog(MsoFileDialogType fileDialogType)
    {
        var dialog = _application?.FileDialog[(MsCore.MsoFileDialogType)(int)fileDialogType];
        return dialog != null ? new MudTools.OfficeInterop.Imps.OfficeFileDialog(dialog) : null;
    }

    public void Activate()
    {
        try
        {
            _application?.Visible = true;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to activate Word application.", ex);
        }
    }


    public int WindowStateValue
    {
        get
        {
            if (_application == null)
                throw new ObjectDisposedException(nameof(_application));
            return _application.WindowState.ConvertToInt();
        }
        set
        {
            if (_application == null)
                throw new ObjectDisposedException(nameof(_application));
            _application.WindowState = value.EnumConvert<MsExcel.XlWindowState>();
        }
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
            var excelwindows = Wn != null ? new ExcelWindow(Wn) : null;
            _windowResize(excelWorkbook, excelwindows);
        }
    }

    private void OnWindowDeactivate(MsExcel.Workbook Wb, MsExcel.Window Wn)
    {
        if (_windowDeactivate != null && Wb != null)
        {
            var excelWorkbook = new ExcelWorkbook(Wb);
            var excelwindows = Wn != null ? new ExcelWindow(Wn) : null;
            _windowDeactivate(excelWorkbook, excelwindows);
        }
    }

    private void OnWindowActivate(MsExcel.Workbook Wb, MsExcel.Window Wn)
    {
        if (_windowActivate != null && Wb != null)
        {
            var excelWorkbook = new ExcelWorkbook(Wb);
            var excelwindows = Wn != null ? new ExcelWindow(Wn) : null;
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
        IExcelComSheet? excelComSheet = null;
        if (sheet != null)
            excelComSheet = Utils.CreateSheetObj(sheet);
        IExcelRange? excelRange = null;
        if (range != null)
            excelRange = new ExcelRange(range);

        if (_sheetChange != null && excelComSheet != null && excelRange != null)
        {
            _sheetChange(excelComSheet, excelRange);
        }
    }

    /// <summary>
    /// 处理SheetActivate事件
    /// </summary>
    /// <param name="sheet">激活的工作表</param>
    private void OnSheetActivate(object sheet)
    {
        IExcelComSheet? excelComSheet = null;
        if (sheet != null)
            excelComSheet = Utils.CreateSheetObj(sheet);
        if (_sheetActivate != null && excelComSheet != null)
        {
            _sheetActivate(excelComSheet);
        }
    }

    /// <summary>
    /// 处理SheetDeactivate事件
    /// </summary>
    /// <param name="sheet">取消激活的工作表</param>
    private void OnSheetDeactivate(object sheet)
    {
        IExcelComSheet? excelComSheet = null;
        if (sheet != null)
            excelComSheet = Utils.CreateSheetObj(sheet);
        if (_sheetDeactivate != null && excelComSheet != null)
        {
            _sheetDeactivate(excelComSheet);
        }
    }

    /// <summary>
    /// 处理SheetSelectionChange事件
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="range">选中的区域</param>
    private void OnSheetSelectionChange(object sheet, MsExcel.Range range)
    {
        IExcelComSheet? excelComSheet = null;
        if (sheet != null)
            excelComSheet = Utils.CreateSheetObj(sheet);
        IExcelRange? excelRange = null;
        if (range != null)
            excelRange = new ExcelRange(range);

        if (_sheetSelectionChange != null && excelComSheet != null && excelRange != null)
        {
            _sheetSelectionChange(excelComSheet, excelRange);
        }
    }

    /// <summary>
    /// 处理SheetBeforeDoubleClick事件
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="range">双击的单元格</param>
    /// <param name="cancel">是否取消默认操作</param>
    private void OnSheetBeforeDoubleClick(object sheet, MsExcel.Range range, ref bool cancel)
    {
        IExcelComSheet? excelComSheet = null;
        if (sheet != null)
            excelComSheet = Utils.CreateSheetObj(sheet);
        IExcelRange? excelRange = null;
        if (range != null)
            excelRange = new ExcelRange(range);

        if (_sheetBeforeDoubleClick != null && excelComSheet != null && excelRange != null)
        {
            _sheetBeforeDoubleClick(excelComSheet, excelRange, ref cancel);
        }
    }

    /// <summary>
    /// 处理SheetBeforeRightClick事件
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="range">右击的单元格</param>
    /// <param name="cancel">是否取消默认操作</param>
    private void OnSheetBeforeRightClick(object sheet, MsExcel.Range range, ref bool cancel)
    {
        IExcelComSheet? excelComSheet = null;
        if (sheet != null)
            excelComSheet = Utils.CreateSheetObj(sheet);
        IExcelRange? excelRange = null;
        if (range != null)
            excelRange = new ExcelRange(range);

        if (_sheetBeforeRightClick != null && excelComSheet != null && excelRange != null)
        {
            _sheetBeforeRightClick(excelComSheet, excelRange, ref cancel);
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


    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_application != null)
            {
                DisConnectEvent();
                // 释放 COM 对象
                try { while (0 < Marshal.ReleaseComObject(_application)) { } } catch { }
                _application = null;
            }
            _disposableList.Dispose();
            _excelRange_ActiveCell?.Dispose();
            _excelRange_ActiveCell = null;
            _excelChart_ActiveChart?.Dispose();
            _excelChart_ActiveChart = null;
            _excelWindow_ActiveWindow?.Dispose();
            _excelWindow_ActiveWindow = null;
            _excelWorkbook_ActiveWorkbook?.Dispose();
            _excelWorkbook_ActiveWorkbook = null;
            _excelAddIns_AddIns?.Dispose();
            _excelAddIns_AddIns = null;
            _excelRange_Cells?.Dispose();
            _excelRange_Cells = null;
            _excelSheets_Charts?.Dispose();
            _excelSheets_Charts = null;
            _excelRange_Columns?.Dispose();
            _excelRange_Columns = null;
            _excelNames_Names?.Dispose();
            _excelNames_Names = null;
            _excelRange_Rows?.Dispose();
            _excelRange_Rows = null;
            _excelSheets_Sheets?.Dispose();
            _excelSheets_Sheets = null;
            _excelWorkbook_ThisWorkbook?.Dispose();
            _excelWorkbook_ThisWorkbook = null;
            _excelWindows_Windows?.Dispose();
            _excelWindows_Windows = null;
            _excelWorkbooks_Workbooks?.Dispose();
            _excelWorkbooks_Workbooks = null;
            _excelWorksheetFunction_WorksheetFunction?.Dispose();
            _excelWorksheetFunction_WorksheetFunction = null;
            _excelSheets_Worksheets?.Dispose();
            _excelSheets_Worksheets = null;
            _excelSheets_Excel4IntlMacroSheets?.Dispose();
            _excelSheets_Excel4IntlMacroSheets = null;
            _excelSheets_Excel4MacroSheets?.Dispose();
            _excelSheets_Excel4MacroSheets = null;
            _excelAutoCorrect_AutoCorrect?.Dispose();
            _excelAutoCorrect_AutoCorrect = null;
            _excelDialogs_Dialogs?.Dispose();
            _excelDialogs_Dialogs = null;
            _officeFileSearch_FileSearch?.Dispose();
            _officeFileSearch_FileSearch = null;
            _excelRecentFiles_RecentFiles?.Dispose();
            _excelRecentFiles_RecentFiles = null;
            _excelODBCErrors_ODBCErrors?.Dispose();
            _excelODBCErrors_ODBCErrors = null;
            _vbeApplication_VBE?.Dispose();
            _vbeApplication_VBE = null;
            _excelOLEDBErrors_OLEDBErrors?.Dispose();
            _excelOLEDBErrors_OLEDBErrors = null;
            _officeCOMAddIns_COMAddIns?.Dispose();
            _officeCOMAddIns_COMAddIns = null;
            _excelCellFormat_FindFormat?.Dispose();
            _excelCellFormat_FindFormat = null;
            _excelCellFormat_ReplaceFormat?.Dispose();
            _excelCellFormat_ReplaceFormat = null;
            _excelWatches_Watches?.Dispose();
            _excelWatches_Watches = null;
            _excelAutoRecover_AutoRecover?.Dispose();
            _excelAutoRecover_AutoRecover = null;
            _excelErrorCheckingOptions_ErrorCheckingOptions?.Dispose();
            _excelErrorCheckingOptions_ErrorCheckingOptions = null;
            _excelSmartTagRecognizers_SmartTagRecognizers?.Dispose();
            _excelSmartTagRecognizers_SmartTagRecognizers = null;
            _officeNewFile_NewWorkbook?.Dispose();
            _officeNewFile_NewWorkbook = null;
            _excelSpellingOptions_SpellingOptions?.Dispose();
            _excelSpellingOptions_SpellingOptions = null;
            _excelSpeech_Speech?.Dispose();
            _excelSpeech_Speech = null;
            _excelRange_ThisCell?.Dispose();
            _excelRange_ThisCell = null;
            _officeAssistance_Assistance?.Dispose();
            _officeAssistance_Assistance = null;
            _excelMultiThreadedCalculation_MultiThreadedCalculation?.Dispose();
            _excelMultiThreadedCalculation_MultiThreadedCalculation = null;
            _excelFileExportConverters_FileExportConverters?.Dispose();
            _excelFileExportConverters_FileExportConverters = null;
            _officeSmartArtLayouts_SmartArtLayouts?.Dispose();
            _officeSmartArtLayouts_SmartArtLayouts = null;
            _officeSmartArtQuickStyles_SmartArtQuickStyles?.Dispose();
            _officeSmartArtQuickStyles_SmartArtQuickStyles = null;
            _officeSmartArtColors_SmartArtColors?.Dispose();
            _officeSmartArtColors_SmartArtColors = null;
            _excelAddIns2_AddIns2?.Dispose();
            _excelAddIns2_AddIns2 = null;
            _excelProtectedViewWindows_ProtectedViewWindows?.Dispose();
            _excelProtectedViewWindows_ProtectedViewWindows = null;
            _excelProtectedViewWindow_ActiveProtectedViewWindow?.Dispose();
            _excelProtectedViewWindow_ActiveProtectedViewWindow = null;
            _excelQuickAnalysis_QuickAnalysis?.Dispose();
            _excelQuickAnalysis_QuickAnalysis = null;
            _officeLanguageSettings_LanguageSettings?.Dispose();
            _officeLanguageSettings_LanguageSettings = null;
            _officeCommandBars_CommandBars?.Dispose();
            _officeCommandBars_CommandBars = null;
            GC.Collect();
        }

        _disposedValue = true;
    }
}
