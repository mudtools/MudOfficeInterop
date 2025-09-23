//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
partial class ExcelWorkbook
{
    private MsExcel.WorkbookEvents_Event _workbookEvents_Event;

    private WorkBookSyncEventHandler _workBookSyncEventHandler;
    private WorkBookSheetChangeEventHandler _workBookSheetChangeEventHandler;
    private WorkbookActivateEventHandler _workbookActivateEventHandler;
    private WorkBookAfterSaveEventHandler _workBookAfterSaveEventHandler;
    private SheetSelectionChangeEventHandler _sheetSelectionChangeEventHandler;
    private WorkbookBeforeCloseEventHandler _workbookBeforeCloseEventHandler;
    private WorkBookNewChartEventHandler _workbookNewChartEventHandler;
    private WorkBookBeforePrintEventHandler _workbookBeforePrintEventHandler;
    private DeactivateEventHandler _deActivateEventHandler;
    private WorkBookNewSheetEventHandler _workbookNewSheetEventHandler;
    private SheetDeactivateEventHandler _workbookSheetDeactivateEventHandler;
    private WorkbookOpenEventHandler _workbookOpenEventHandler;
    private WorkBookPivotTableCloseConnectionEventHandler _workBookPivotTableCloseConnectionEventHandler;
    private WorkBookPivotTableOpenConnectionEventHandler _workBookPivotTableOpenConnectionEventHandler;
    private WorkBookBeforeRowsetCompleteEventHandler _workBookBeforeRowsetCompleteEventHandler;
    private SheetBeforeDeleteEventHandler _sheetBeforeDeleteEventHandler;
    private SheetCalculateEventHandler _sheetCalculateEventHandler;
    private WindowActivateEventHandler _windowActivateEventHandler;
    private WindowDeactivateEventHandler _windowDeactivateEventHandler;
    private WindowResizeEventHandler _windowResizeEventHandler;
    private WorkBookSheetPivotTableChangeEventHandler _workbookSheetPivotTableChangeEventHandler;

    /// <summary>
    /// 初始化Excel事件处理
    /// </summary>
    private void InitializeEvents()
    {
        if (_workbookEvents_Event == null) return;

        _workbookEvents_Event.Sync += _workbookEvents_Event_Sync;
        _workbookEvents_Event.Activate += _workbookEvents_Event_Activate;
        _workbookEvents_Event.AfterSave += _workbookEvents_Event_AfterSave;
        _workbookEvents_Event.BeforeClose += _workbookEvents_Event_BeforeClose;
        _workbookEvents_Event.NewChart += _workbookEvents_Event_NewChart;
        _workbookEvents_Event.BeforePrint += _workbookEvents_Event_BeforePrint;
        _workbookEvents_Event.Deactivate += _workbookEvents_Event_Deactivate;
        _workbookEvents_Event.Open += _workbookEvents_Event_Open;
        _workbookEvents_Event.PivotTableCloseConnection += _workbookEvents_Event_PivotTableCloseConnection;
        _workbookEvents_Event.PivotTableOpenConnection += _workbookEvents_Event_PivotTableOpenConnection;
        _workbookEvents_Event.RowsetComplete += _workbookEvents_Event_RowsetComplete;

        _workbookEvents_Event.NewSheet += _workbookEvents_Event_NewSheet;
        _workbookEvents_Event.SheetChange += _workbookEvents_Event_SheetChange;
        _workbookEvents_Event.SheetSelectionChange += _workbookEvents_Event_SheetSelectionChange;
        _workbookEvents_Event.SheetDeactivate += _workbookEvents_Event_SheetDeactivate;
        _workbookEvents_Event.SheetBeforeDelete += _workbookEvents_Event_SheetBeforeDelete;
        _workbookEvents_Event.SheetCalculate += _workbookEvents_Event_SheetCalculate;
        _workbookEvents_Event.SheetPivotTableChangeSync += _workbookEvents_Event_SheetPivotTableChangeSync;

        _workbookEvents_Event.WindowActivate += _workbookEvents_Event_WindowActivate;
        _workbookEvents_Event.WindowDeactivate += _workbookEvents_Event_WindowDeactivate;
        _workbookEvents_Event.WindowResize += _workbookEvents_Event_WindowResize;
    }

    private void DisConnectEvent()
    {
        if (_workbookEvents_Event == null)
            return;

        _workbookEvents_Event.Sync -= _workbookEvents_Event_Sync;
        _workbookEvents_Event.Activate -= _workbookEvents_Event_Activate;
        _workbookEvents_Event.AfterSave -= _workbookEvents_Event_AfterSave;
        _workbookEvents_Event.BeforeClose -= _workbookEvents_Event_BeforeClose;
        _workbookEvents_Event.NewChart -= _workbookEvents_Event_NewChart;
        _workbookEvents_Event.BeforePrint -= _workbookEvents_Event_BeforePrint;
        _workbookEvents_Event.Deactivate -= _workbookEvents_Event_Deactivate;
        _workbookEvents_Event.Open -= _workbookEvents_Event_Open;
        _workbookEvents_Event.PivotTableCloseConnection -= _workbookEvents_Event_PivotTableCloseConnection;
        _workbookEvents_Event.PivotTableOpenConnection -= _workbookEvents_Event_PivotTableOpenConnection;
        _workbookEvents_Event.RowsetComplete -= _workbookEvents_Event_RowsetComplete;

        _workbookEvents_Event.NewSheet -= _workbookEvents_Event_NewSheet;
        _workbookEvents_Event.SheetChange -= _workbookEvents_Event_SheetChange;
        _workbookEvents_Event.SheetSelectionChange -= _workbookEvents_Event_SheetSelectionChange;
        _workbookEvents_Event.SheetDeactivate -= _workbookEvents_Event_SheetDeactivate;
        _workbookEvents_Event.SheetBeforeDelete -= _workbookEvents_Event_SheetBeforeDelete;
        _workbookEvents_Event.SheetCalculate -= _workbookEvents_Event_SheetCalculate;
        _workbookEvents_Event.SheetPivotTableChangeSync -= _workbookEvents_Event_SheetPivotTableChangeSync;

        _workbookEvents_Event.WindowActivate -= _workbookEvents_Event_WindowActivate;
        _workbookEvents_Event.WindowDeactivate -= _workbookEvents_Event_WindowDeactivate;
        _workbookEvents_Event.WindowResize -= _workbookEvents_Event_WindowResize;
    }

    private void _workbookEvents_Event_WindowResize(MsExcel.Window Wn)
    {
        var workSheet = Wn as MsExcel.Window;
        using var window = new ExcelWindow(workSheet);
        _windowResizeEventHandler?.Invoke(this, window);
    }

    private void _workbookEvents_Event_WindowDeactivate(MsExcel.Window Wn)
    {
        var workSheet = Wn as MsExcel.Window;
        using var window = new ExcelWindow(workSheet);
        _windowDeactivateEventHandler?.Invoke(this, window);
    }

    private void _workbookEvents_Event_WindowActivate(MsExcel.Window Wn)
    {
        var workSheet = Wn as MsExcel.Window;
        using var window = new ExcelWindow(workSheet);
        _windowActivateEventHandler?.Invoke(this, window);
    }

    private void _workbookEvents_Event_SheetPivotTableChangeSync(object Sh, MsExcel.PivotTable Target)
    {
        var workSheet = Sh as MsExcel.Worksheet;
        using var sheet = new ExcelWorksheet(workSheet);
        using var pivotTable = new ExcelPivotTable(Target);

        _workbookSheetPivotTableChangeEventHandler?.Invoke(sheet, pivotTable);
    }

    private void _workbookEvents_Event_SheetCalculate(object Sh)
    {
        var workSheet = Sh as MsExcel.Worksheet;
        using var sheet = new ExcelWorksheet(workSheet);
        _sheetCalculateEventHandler?.Invoke(sheet);
    }

    private void _workbookEvents_Event_SheetBeforeDelete(object Sh)
    {
        var workSheet = Sh as MsExcel.Worksheet;
        using var sheet = new ExcelWorksheet(workSheet);
        _sheetBeforeDeleteEventHandler?.Invoke(sheet);
    }

    private void _workbookEvents_Event_SheetDeactivate(object Sh)
    {
        var workSheet = Sh as MsExcel.Worksheet;
        using var sheet = new ExcelWorksheet(workSheet);
        _workbookSheetDeactivateEventHandler?.Invoke(sheet);
    }

    private void _workbookEvents_Event_RowsetComplete(string Description, string Sheet, bool Success)
    {
        _workBookBeforeRowsetCompleteEventHandler?.Invoke(Description, Sheet, Success);
    }

    private void _workbookEvents_Event_PivotTableOpenConnection(MsExcel.PivotTable Target)
    {
        using var pivotTable = new ExcelPivotTable(Target);
        _workBookPivotTableOpenConnectionEventHandler?.Invoke(pivotTable);
    }

    private void _workbookEvents_Event_PivotTableCloseConnection(MsExcel.PivotTable Target)
    {
        using var pivotTable = new ExcelPivotTable(Target);
        _workBookPivotTableCloseConnectionEventHandler?.Invoke(pivotTable);
    }

    private void _workbookEvents_Event_Open()
    {
        _workbookOpenEventHandler?.Invoke(this);
    }

    private void _workbookEvents_Event_NewSheet(object Sh)
    {
        var workSheet = Sh as MsExcel.Worksheet;
        using var sheet = new ExcelWorksheet(workSheet);
        _workbookNewSheetEventHandler?.Invoke(sheet);
    }

    private void _workbookEvents_Event_Deactivate()
    {
        _deActivateEventHandler?.Invoke();
    }

    private void _workbookEvents_Event_BeforePrint(ref bool cancel)
    {
        _workbookBeforePrintEventHandler?.Invoke(cancel);
    }

    private void _workbookEvents_Event_NewChart(MsExcel.Chart Ch)
    {
        using IExcelChart chart = new ExcelChart(Ch);
        _workbookNewChartEventHandler?.Invoke(chart);
    }

    private void _workbookEvents_Event_BeforeClose(ref bool Cancel)
    {
        _workbookBeforeCloseEventHandler?.Invoke(this, ref Cancel);
    }

    private void _workbookEvents_Event_SheetSelectionChange(object Sh, MsExcel.Range Target)
    {
        var workSheet = Sh as MsExcel.Worksheet;
        using var sheet = new ExcelWorksheet(workSheet);
        using var range = new ExcelRange(Target);
        _sheetSelectionChangeEventHandler?.Invoke(sheet, range);
    }

    private void _workbookEvents_Event_AfterSave(bool Success)
    {
        _workBookAfterSaveEventHandler?.Invoke(Success);
    }

    private void _workbookEvents_Event_Activate()
    {
        _workbookActivateEventHandler?.Invoke(this);
    }

    private void _workbookEvents_Event_SheetChange(object Sh, MsExcel.Range Target)
    {
        using var range = new ExcelRange(Target);
        _workBookSheetChangeEventHandler?.Invoke(Sh, range);
    }

    private void _workbookEvents_Event_Sync(MsCore.MsoSyncEventType SyncEventType)
    {
        _workBookSyncEventHandler?.Invoke((MsoSyncEventType)(int)SyncEventType);
    }

    public event WorkBookSheetPivotTableChangeEventHandler WorkBookSheetPivotTableChange
    {
        add { _workbookSheetPivotTableChangeEventHandler = value; }
        remove { _workbookSheetPivotTableChangeEventHandler -= value; }
    }

    public event WindowResizeEventHandler WindowResize
    {
        add { _windowResizeEventHandler = value; }
        remove { _windowResizeEventHandler -= value; }
    }

    public event WindowDeactivateEventHandler WindowDeActivate
    {
        add { _windowDeactivateEventHandler = value; }
        remove { _windowDeactivateEventHandler -= value; }
    }

    public event WindowActivateEventHandler WindowActivate
    {
        add { _windowActivateEventHandler = value; }
        remove { _windowActivateEventHandler -= value; }
    }

    public event SheetCalculateEventHandler Calculate
    {
        add { _sheetCalculateEventHandler = value; }
        remove { _sheetCalculateEventHandler -= value; }
    }

    public event SheetBeforeDeleteEventHandler SheetBeforeDelete
    {
        add { _sheetBeforeDeleteEventHandler = value; }
        remove { _sheetBeforeDeleteEventHandler -= value; }
    }

    public event SheetDeactivateEventHandler SheetDeactivate
    {
        add { _workbookSheetDeactivateEventHandler = value; }
        remove { _workbookSheetDeactivateEventHandler -= value; }
    }

    public event WorkBookBeforeRowsetCompleteEventHandler OnBeforeRowsetComplete
    {
        add { _workBookBeforeRowsetCompleteEventHandler = value; }
        remove { _workBookBeforeRowsetCompleteEventHandler -= value; }
    }

    public event WorkBookPivotTableOpenConnectionEventHandler PivotTableOpenConnection
    {
        add { _workBookPivotTableOpenConnectionEventHandler += value; }
        remove { _workBookPivotTableOpenConnectionEventHandler -= value; }
    }

    public event WorkBookPivotTableCloseConnectionEventHandler PivotTableCloseConnection
    {
        add { _workBookPivotTableCloseConnectionEventHandler += value; }
        remove { _workBookPivotTableCloseConnectionEventHandler -= value; }
    }

    public event WorkbookOpenEventHandler Open
    {
        add { _workbookOpenEventHandler += value; }
        remove { _workbookOpenEventHandler -= value; }
    }

    public event WorkBookNewSheetEventHandler NewSheet
    {
        add { _workbookNewSheetEventHandler += value; }
        remove { _workbookNewSheetEventHandler -= value; }
    }

    public event DeactivateEventHandler Deactivate
    {
        add { _deActivateEventHandler += value; }
        remove { _deActivateEventHandler -= value; }
    }

    public event WorkBookBeforePrintEventHandler BeforePrint
    {
        add { _workbookBeforePrintEventHandler += value; }
        remove { _workbookBeforePrintEventHandler -= value; }
    }

    public event WorkBookNewChartEventHandler NewChart
    {
        add { _workbookNewChartEventHandler += value; }
        remove { _workbookNewChartEventHandler -= value; }
    }

    public event WorkbookBeforeCloseEventHandler BeforeClose
    {
        add { _workbookBeforeCloseEventHandler += value; }
        remove { _workbookBeforeCloseEventHandler -= value; }
    }

    public event SheetSelectionChangeEventHandler SheetSelectionChange
    {
        add { _sheetSelectionChangeEventHandler += value; }
        remove { _sheetSelectionChangeEventHandler -= value; }
    }

    public event WorkBookAfterSaveEventHandler AfterSave
    {
        add { _workBookAfterSaveEventHandler += value; }
        remove { _workBookAfterSaveEventHandler -= value; }
    }

    public event WorkbookActivateEventHandler WorkbookActivate
    {
        add { _workbookActivateEventHandler += value; }
        remove { _workbookActivateEventHandler -= value; }
    }

    public event WorkBookSheetChangeEventHandler SheetChange
    {
        add { _workBookSheetChangeEventHandler += value; }
        remove { _workBookSheetChangeEventHandler -= value; }
    }
    public event WorkBookSyncEventHandler Sync
    {
        add { _workBookSyncEventHandler += value; }
        remove { _workBookSyncEventHandler -= value; }
    }
}
