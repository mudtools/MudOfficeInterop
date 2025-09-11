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
