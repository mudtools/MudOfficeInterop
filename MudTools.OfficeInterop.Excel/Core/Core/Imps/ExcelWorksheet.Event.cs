//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

partial class ExcelWorksheet
{
    /// <summary>
    /// 初始化 ExcelWorksheet 实例
    /// </summary>
    /// <param name="worksheet">底层的 COM Worksheet 对象</param>
    internal ExcelWorksheet(MsExcel.Worksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _docEvents_Event = worksheet;
        InitializeEvents();
        _disposedValue = false;
    }

    public string? ParentName
    {
        get
        {
            if (_worksheet?.Parent == null)
            {
                return null;
            }
            if (_worksheet.Parent is MsExcel.Workbook workbook)
            {
                return workbook.Name;
            }
            if (_worksheet.Parent is MsExcel.Worksheet worksheet)
            {
                return worksheet.Name;
            }
            return null;
        }
    }

    public IExcelWorkbook? ParentWorkbook
    {
        get
        {
            if (_worksheet?.Parent == null)
            {
                return null;
            }
            if (_worksheet.Parent is MsExcel.Workbook workbook)
            {
                return new ExcelWorkbook(workbook);
            }
            return null;
        }
    }

    /// <summary>
    /// 获取工作表是否被保护
    /// </summary>
    public bool IsProtected => _worksheet != null && _worksheet.ProtectContents;

    public bool IsVisible
    {
        get => _worksheet != null && _worksheet.Visible == MsExcel.XlSheetVisibility.xlSheetVisible;
        set
        {
            if (_worksheet != null)
                _worksheet.Visible = value ? (MsExcel.XlSheetVisibility.xlSheetVisible) : (MsExcel.XlSheetVisibility.xlSheetHidden);
        }
    }

    /// <summary>
    /// 复制工作表
    /// </summary>
    /// <param name="before">复制到指定工作表之前</param>
    /// <param name="after">复制到指定工作表之后</param>
    public void Copy(IExcelComSheet? before = null, IExcelComSheet? after = null)
    {
        if (_worksheet == null) return;

        _worksheet.Copy(
            before is ExcelWorksheet beforeSheet ? beforeSheet._worksheet : System.Type.Missing,
            after is ExcelWorksheet afterSheet ? afterSheet._worksheet : System.Type.Missing
        );
    }

    /// <summary>
    /// 移动工作表
    /// </summary>
    /// <param name="before">移动到指定工作表之前</param>
    /// <param name="after">移动到指定工作表之后</param>
    public void Move(IExcelComSheet? before = null, IExcelComSheet? after = null)
    {
        if (_worksheet == null) return;

        _worksheet.Move(
            before is ExcelWorksheet beforeSheet ? beforeSheet._worksheet : System.Type.Missing,
            after is ExcelWorksheet afterSheet ? afterSheet._worksheet : System.Type.Missing
        );
    }

    #region 事件字段

    private MsExcel.DocEvents_Event _docEvents_Event;

    /// <summary>
    /// Change事件
    /// </summary>
    private event ChangeEventHandler _change;

    /// <summary>
    /// SelectionChange事件
    /// </summary>
    private event SelectionChangeEventHandler _selectionChange;

    /// <summary>
    /// Activate事件
    /// </summary>
    private event ActivateEventHandler _sheetActivate;

    /// <summary>
    /// Deactivate事件
    /// </summary>
    private event DeactivateEventHandler _sheetDeactivate;

    /// <summary>
    /// BeforeDoubleClick事件
    /// </summary>
    private event BeforeDoubleClickEventHandler _beforeDoubleClick;

    /// <summary>
    /// BeforeRightClick事件
    /// </summary>
    private event BeforeRightClickEventHandler _beforeRightClick;

    /// <summary>
    /// Calculate事件
    /// </summary>
    private event CalculateEventHandler _sheetCalculate;

    private event BeforeDeleteEventHandler _sheetBeforeDelete;

    private event PivotTableChangeSyncEventHandler _sheetPivotTableChangeSync;

    #endregion

    /// <summary>
    /// 初始化Excel事件处理
    /// </summary>
    private void InitializeEvents()
    {
        _docEvents_Event.Activate += _docEvents_Event_Activate;
        _docEvents_Event.BeforeDelete += _docEvents_Event_BeforeDelete;
        _docEvents_Event.BeforeDoubleClick += _docEvents_Event_BeforeDoubleClick;
        _docEvents_Event.BeforeRightClick += _docEvents_Event_BeforeRightClick;
        _docEvents_Event.Change += _docEvents_Event_Change;
        _docEvents_Event.Calculate += _docEvents_Event_Calculate;
        _docEvents_Event.PivotTableChangeSync += _docEvents_Event_PivotTableChangeSync;
        _docEvents_Event.Deactivate += _docEvents_Event_Deactivate;
        _docEvents_Event.SelectionChange += _docEvents_Event_SelectionChange;
    }

    private void DisConnectEvent()
    {
        if (_docEvents_Event == null)
            return;

        _docEvents_Event.Activate -= _docEvents_Event_Activate;
        _docEvents_Event.BeforeDelete -= _docEvents_Event_BeforeDelete;
        _docEvents_Event.BeforeDoubleClick -= _docEvents_Event_BeforeDoubleClick;
        _docEvents_Event.BeforeRightClick -= _docEvents_Event_BeforeRightClick;
        _docEvents_Event.Change -= _docEvents_Event_Change;
        _docEvents_Event.Calculate -= _docEvents_Event_Calculate;
        _docEvents_Event.PivotTableChangeSync -= _docEvents_Event_PivotTableChangeSync;
        _docEvents_Event.Deactivate -= _docEvents_Event_Deactivate;
        _docEvents_Event.SelectionChange -= _docEvents_Event_SelectionChange;
    }

    private void _docEvents_Event_SelectionChange(MsExcel.Range target)
    {
        var range = new ExcelRange(target);
        _selectionChange?.Invoke(range);
    }

    private void _docEvents_Event_Deactivate()
    {
        _sheetDeactivate?.Invoke();
    }

    private void _docEvents_Event_PivotTableChangeSync(MsExcel.PivotTable target)
    {
        var pivotTable = new ExcelPivotTable(target);
        _sheetPivotTableChangeSync?.Invoke(pivotTable);
    }

    private void _docEvents_Event_Calculate()
    {
        _sheetCalculate?.Invoke();
    }

    private void _docEvents_Event_Change(MsExcel.Range target)
    {
        var range = new ExcelRange(target);
        _change?.Invoke(range);
    }

    private void _docEvents_Event_BeforeRightClick(MsExcel.Range target, ref bool cancel)
    {
        var range = new ExcelRange(target);
        _beforeRightClick?.Invoke(range, ref cancel);
    }

    private void _docEvents_Event_BeforeDoubleClick(MsExcel.Range target, ref bool cancel)
    {
        var range = new ExcelRange(target);
        _beforeDoubleClick?.Invoke(range, ref cancel);
    }

    private void _docEvents_Event_BeforeDelete()
    {
        _sheetBeforeDelete?.Invoke();
    }

    private void _docEvents_Event_Activate()
    {
        _sheetActivate?.Invoke();
    }

    #region 事件实现

    /// <summary>
    /// 当工作表内容发生改变时触发
    /// </summary>
    public event ChangeEventHandler Change
    {
        add { _change += value; }
        remove { _change -= value; }
    }

    /// <summary>
    /// 当工作表选择区域发生改变时触发
    /// </summary>
    public event SelectionChangeEventHandler SelectionChange
    {
        add { _selectionChange += value; }
        remove { _selectionChange -= value; }
    }

    /// <summary>
    /// 当工作表被激活时触发
    /// </summary>
    public event ActivateEventHandler SheetActivate
    {
        add { _sheetActivate += value; }
        remove { _sheetActivate -= value; }
    }

    /// <summary>
    /// 当工作表被取消激活时触发
    /// </summary>
    public event DeactivateEventHandler SheetDeactivate
    {
        add { _sheetDeactivate += value; }
        remove { _sheetDeactivate -= value; }
    }

    /// <summary>
    /// 当工作表被双击时触发
    /// </summary>
    public event BeforeDoubleClickEventHandler BeforeDoubleClick
    {
        add { _beforeDoubleClick += value; }
        remove { _beforeDoubleClick -= value; }
    }

    /// <summary>
    /// 当工作表被右键单击时触发
    /// </summary>
    public event BeforeRightClickEventHandler BeforeRightClick
    {
        add { _beforeRightClick += value; }
        remove { _beforeRightClick -= value; }
    }

    /// <summary>
    /// 当工作表计算完成后触发
    /// </summary>
    public event CalculateEventHandler SheetCalculate
    {
        add { _sheetCalculate += value; }
        remove { _sheetCalculate -= value; }
    }

    public event BeforeDeleteEventHandler BeforeDelete
    {
        add { _sheetBeforeDelete += value; }
        remove { _sheetBeforeDelete -= value; }
    }

    public event PivotTableChangeSyncEventHandler PivotTableChangeSync
    {
        add { _sheetPivotTableChangeSync += value; }
        remove { _sheetPivotTableChangeSync -= value; }
    }
    #endregion

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            DisConnectEvent();

            if (_worksheet != null)
            {
                Marshal.ReleaseComObject(_worksheet);
                _worksheet = null;
            }
            _excelChart_NextChart?.Dispose();
            _excelChart_NextChart = null;
            _excelRange_NextRange?.Dispose();
            _excelRange_NextRange = null;
            _excelWorksheet_NextWorksheet?.Dispose();
            _excelWorksheet_NextWorksheet = null;
            _excelChart_PreviousChart?.Dispose();
            _excelChart_PreviousChart = null;
            _excelRange_PreviousRange?.Dispose();
            _excelRange_PreviousRange = null;
            _excelWorksheet_PreviousWorksheet?.Dispose();
            _excelWorksheet_PreviousWorksheet = null;
            _excelRange_Cells?.Dispose();
            _excelRange_Cells = null;
            _excelRange_CircularReference?.Dispose();
            _excelRange_CircularReference = null;
            _excelRange_Columns?.Dispose();
            _excelRange_Columns = null;
            _excelNames_Names?.Dispose();
            _excelNames_Names = null;
            _excelOutline_Outline?.Dispose();
            _excelOutline_Outline = null;
            _excelRange_Rows?.Dispose();
            _excelRange_Rows = null;
            _excelRange_UsedRange?.Dispose();
            _excelRange_UsedRange = null;
            _excelHPageBreaks_HPageBreaks?.Dispose();
            _excelHPageBreaks_HPageBreaks = null;
            _excelVPageBreaks_VPageBreaks?.Dispose();
            _excelVPageBreaks_VPageBreaks = null;
            _excelQueryTables_QueryTables?.Dispose();
            _excelQueryTables_QueryTables = null;
            _excelComments_Comments?.Dispose();
            _excelComments_Comments = null;
            _excelAutoFilter_AutoFilter?.Dispose();
            _excelAutoFilter_AutoFilter = null;
            _excelTab_Tab?.Dispose();
            _excelTab_Tab = null;
            _officeMsoEnvelope_MailEnvelope?.Dispose();
            _officeMsoEnvelope_MailEnvelope = null;
            _excelCustomProperties_CustomProperties?.Dispose();
            _excelCustomProperties_CustomProperties = null;
            _excelProtection_Protection?.Dispose();
            _excelProtection_Protection = null;
            _excelListObjects_ListObjects?.Dispose();
            _excelListObjects_ListObjects = null;
            _excelSort_Sort?.Dispose();
            _excelSort_Sort = null;
        }

        _disposedValue = true;
    }

}
