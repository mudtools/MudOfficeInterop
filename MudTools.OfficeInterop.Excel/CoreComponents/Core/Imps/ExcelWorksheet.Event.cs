//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
partial class ExcelWorksheet
{
    private MsExcel.DocEvents_Event _docEvents_Event;

    #region 事件字段

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
}
