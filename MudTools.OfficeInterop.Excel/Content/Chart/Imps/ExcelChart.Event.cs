//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
partial class ExcelChart
{
    private MsExcel.ChartEvents_Event? _chartEvents_Event;

    private ChartActivateEventHandler? _chartActivateEventHandler;
    private ChartDeactivateEventHandler? _chartDeactivateEventHandler;
    private ChartSelectEventHandler? _chartSelectEventHandler;
    private ChartBeforeDoubleClickEventHandler? _chartBeforeDoubleClickEventHandler;
    private ChartBeforeRightClickEventHandler? _chartBeforeRightClickEventHandler;
    private ChartSeriesChangeEventHandler? _chartSeriesChangeEventHandler;

    /// <summary>
    /// 初始化 ChartEvents_Event 的所有事件监听器
    /// </summary>
    internal void InitializeEvents()
    {
        if (_chartEvents_Event == null) return;

        _chartEvents_Event.Activate += _chartEvents_Event_Activate;
        _chartEvents_Event.Deactivate += _chartEvents_Event_Deactivate;
        _chartEvents_Event.Select += _chartEvents_Event_Select;
        _chartEvents_Event.BeforeDoubleClick += _chartEvents_Event_BeforeDoubleClick;
        _chartEvents_Event.BeforeRightClick += _chartEvents_Event_BeforeRightClick;
        _chartEvents_Event.SeriesChange += _chartEvents_Event_SeriesChange;
    }

    /// <summary>
    /// 取消所有事件绑定，防止内存泄漏
    /// </summary>
    internal void DisconnectEvents()
    {
        if (_chartEvents_Event == null) return;

        _chartEvents_Event.Activate -= _chartEvents_Event_Activate;
        _chartEvents_Event.Deactivate -= _chartEvents_Event_Deactivate;
        _chartEvents_Event.Select -= _chartEvents_Event_Select;
        _chartEvents_Event.BeforeDoubleClick -= _chartEvents_Event_BeforeDoubleClick;
        _chartEvents_Event.BeforeRightClick -= _chartEvents_Event_BeforeRightClick;
        _chartEvents_Event.SeriesChange -= _chartEvents_Event_SeriesChange;
    }

    #region 私有事件处理器（与 ChartEvents_Event 成员一一对应）

    /// <summary>
    /// 当图表被激活时触发（例如用户点击图表区域）
    /// </summary>
    private void _chartEvents_Event_Activate()
    {
        _chartActivateEventHandler?.Invoke(this);
    }

    /// <summary>
    /// 当图表失去焦点（被其他对象选中）时触发
    /// </summary>
    private void _chartEvents_Event_Deactivate()
    {
        _chartDeactivateEventHandler?.Invoke(this);
    }

    /// <summary>
    /// 当用户在图表上选择某个元素时触发（如数据点、图例、坐标轴等）
    /// 参数说明：
    ///   ElementID: MsoChartElementType 枚举值（int），标识选中的元素类型
    ///   SeriesIndex: 系列索引（int），若选中的是系列或数据点，则为该系列编号；否则为 Missing.Value 或 0
    ///   PointIndex: 数据点索引（int），若选中的是单个数据点，则为该点编号；否则为 Missing.Value 或 0
    /// 注意：若选中的是图例，则 SeriesIndex 和 PointIndex 均为 Missing.Value
    /// </summary>
    private void _chartEvents_Event_Select(int ElementID, int SeriesIndex, int PointIndex)
    {
        _chartSelectEventHandler?.Invoke(this, (MsoChartElementType)ElementID, SeriesIndex, PointIndex);
    }

    /// <summary>
    /// 用户双击图表前触发，可取消默认行为（如打开格式对话框）
    /// </summary>
    private void _chartEvents_Event_BeforeDoubleClick(int ElementID, int SeriesIndex, int PointIndex, ref bool Cancel)
    {
        _chartBeforeDoubleClickEventHandler?.Invoke((MsoChartElementType)ElementID, SeriesIndex, PointIndex, ref Cancel);
    }

    /// <summary>
    /// 用户右键单击图表前触发，可取消默认上下文菜单显示
    /// </summary>
    private void _chartEvents_Event_BeforeRightClick(ref bool Cancel)
    {
        _chartBeforeRightClickEventHandler?.Invoke(ref Cancel);
    }

    /// <summary>
    /// 当图表中的任一数据系列发生更改时触发（如修改数据源、添加/删除点、更新值）
    /// 参数说明：
    ///   SeriesIndex: 发生变化的系列索引（int）
    ///   PointIndex: 若仅单个数据点变化，则为该点索引；否则为 Missing.Value
    /// </summary>
    private void _chartEvents_Event_SeriesChange(int SeriesIndex, int PointIndex)
    {
        _chartSeriesChangeEventHandler?.Invoke(this, SeriesIndex, PointIndex);
    }

    #endregion

    #region 公共事件
    /// <summary>
    /// 图表被激活时触发
    /// </summary>
    public event ChartActivateEventHandler ChartActivate
    {
        add { _chartActivateEventHandler += value; }
        remove { _chartActivateEventHandler -= value; }
    }

    /// <summary>
    /// 图表失活时触发
    /// </summary>
    public event ChartDeactivateEventHandler Deactivate
    {
        add { _chartDeactivateEventHandler += value; }
        remove { _chartDeactivateEventHandler -= value; }
    }

    /// <summary>
    /// 用户在图表上选择任意元素时触发（如点击数据点、图例、标题等）
    /// </summary>
    public event ChartSelectEventHandler ChartSelect
    {
        add { _chartSelectEventHandler += value; }
        remove { _chartSelectEventHandler -= value; }
    }

    /// <summary>
    /// 用户双击图表前触发，设置 Cancel = true 可阻止默认行为（如打开“设置数据系列”对话框）
    /// </summary>
    public event ChartBeforeDoubleClickEventHandler BeforeDoubleClick
    {
        add { _chartBeforeDoubleClickEventHandler += value; }
        remove { _chartBeforeDoubleClickEventHandler -= value; }
    }

    /// <summary>
    /// 用户右键单击图表前触发，设置 Cancel = true 可阻止弹出上下文菜单
    /// </summary>
    public event ChartBeforeRightClickEventHandler BeforeRightClick
    {
        add { _chartBeforeRightClickEventHandler += value; }
        remove { _chartBeforeRightClickEventHandler -= value; }
    }

    /// <summary>
    /// 当图表中的数据系列发生变化时触发（例如单元格数据更新导致图表重绘）
    /// </summary>
    public event ChartSeriesChangeEventHandler SeriesChange
    {
        add { _chartSeriesChangeEventHandler += value; }
        remove { _chartSeriesChangeEventHandler -= value; }
    }
    #endregion
}
