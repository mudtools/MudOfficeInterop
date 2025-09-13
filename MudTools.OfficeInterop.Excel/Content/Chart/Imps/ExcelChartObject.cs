//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel ChartObject 对象的二次封装实现类
/// 负责对 Microsoft.Office.Interop.Excel.ChartObject 对象的安全访问和资源管理
/// </summary>
internal class ExcelChartObject : IExcelChartObject
{
    /// <summary>
    /// 底层的 COM ChartObject 对象
    /// </summary>
    private MsExcel.ChartObject? _chartObject;

    /// <summary>
    /// 底层的形状对象
    /// </summary>
    private MsExcel.ShapeRange? _shapeRange;

    /// <summary>
    /// 底层的图表对象
    /// </summary>
    private MsExcel.Chart? _chart;

    /// <summary>
    /// 标记对象是否已被释放
    /// </summary>
    private bool _disposedValue;

    #region 构造函数和释放

    /// <summary>
    /// 初始化 ExcelChartObject 实例
    /// </summary>
    /// <param name="chartObject">底层的 COM ChartObject 对象</param>
    internal ExcelChartObject(MsExcel.ChartObject chartObject)
    {
        _chartObject = chartObject ?? throw new ArgumentNullException(nameof(chartObject));
        _shapeRange = chartObject.ShapeRange;
        _chart = chartObject.Chart;
        _disposedValue = false;
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
                // 释放图表对象
                _excelChart?.Dispose();

                // 释放形状对象
                _excelShape?.Dispose();

                _pageSetup?.Dispose();

                // 释放底层COM对象
                if (_chart != null)
                    Marshal.ReleaseComObject(_chart);

                if (_shapeRange != null)
                    Marshal.ReleaseComObject(_shapeRange);

                if (_chartObject != null)
                    Marshal.ReleaseComObject(_chartObject);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _excelChart = null;
            _excelShape = null;
            _chart = null;
        }

        _disposedValue = true;
    }

    /// <summary>
    /// 实现 IDisposable 接口的释放方法
    /// </summary>
    public void Dispose() => Dispose(true);

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取图表所在的 Excel Application 对象
    /// </summary>
    public IExcelApplication Application => new ExcelApplication(_chartObject.Application);

    /// <summary>
    /// 获取图表内容是否被保护
    /// </summary>
    public bool IsProtected => _chart.ProtectContents;

    /// <summary>
    /// 页面设置对象缓存
    /// </summary>
    private IExcelPageSetup _pageSetup;

    /// <summary>
    /// 获取工作表的页面设置对象
    /// </summary>
    public IExcelPageSetup PageSetup => _pageSetup ?? (_pageSetup = new ExcelPageSetup(_chart?.PageSetup));


    /// <summary>
    /// 获取或设置图表对象的名称
    /// </summary>
    public string Name
    {
        get => _chartObject?.Name?.ToString();
        set
        {
            if (_chartObject != null && value != null)
                _chartObject.Name = value;
        }
    }

    public bool ProtectContents => _chart.ProtectContents;

    /// <summary>
    /// 获取图表对象的索引位置
    /// </summary>
    public int Index => _chartObject?.Index ?? 0;


    public XlSheetVisibility Visible
    {
        get => _chartObject != null && _chartObject.Visible ? XlSheetVisibility.xlSheetVisible : XlSheetVisibility.xlSheetHidden;
        set
        {
            if (_chartObject != null)
                _chartObject.Visible = value == XlSheetVisibility.xlSheetVisible;
        }
    }

    public bool IsVisible
    {
        get => _chartObject != null && _chartObject.Visible;
        set
        {
            if (_chartObject != null)
                _chartObject.Visible = value;
        }
    }

    /// <summary>
    /// 获取图表对象所在的父对象
    /// </summary>
    public object? Parent
    {
        get
        {
            if (_chartObject?.Parent == null)
            {
                return null;
            }
            if (_chartObject.Parent is MsExcel.ChartObjects chartObjs)
            {
                return new ExcelChartObjects(chartObjs);
            }
            if (_chartObject.Parent is MsExcel.Workbook workbook)
            {
                return new ExcelWorkbook(workbook);
            }
            if (_chartObject.Parent is MsExcel.Worksheet worksheet)
            {
                return new ExcelWorksheet(worksheet);
            }
            return null;
        }
    }

    public string? ParentName
    {
        get
        {
            if (_chartObject?.Parent == null)
            {
                return null;
            }
            if (_chartObject.Parent is MsExcel.ChartObjects chartObjs)
            {
                return "";
            }
            if (_chartObject.Parent is MsExcel.Workbook workbook)
            {
                return workbook.Name;
            }
            if (_chartObject.Parent is MsExcel.Worksheet worksheet)
            {
                return worksheet.Name;
            }
            return null;
        }
    }

    /// <summary>
    /// 形状对象缓存
    /// </summary>
    private IExcelShapeRange _excelShape;

    /// <summary>
    /// 获取图表对象的底层形状对象
    /// </summary>
    public IExcelShapeRange ShapeRange => _excelShape ??= new ExcelShapeRange(_shapeRange);

    #endregion

    #region 位置和大小

    /// <summary>
    /// 获取或设置图表对象的左边距
    /// </summary>
    public double Left
    {
        get => _chartObject?.Left ?? 0;
        set
        {
            if (_chartObject != null)
                _chartObject.Left = value;
        }
    }

    /// <summary>
    /// 获取或设置图表对象的顶边距
    /// </summary>
    public double Top
    {
        get => _chartObject?.Top ?? 0;
        set
        {
            if (_chartObject != null)
                _chartObject.Top = value;
        }
    }

    /// <summary>
    /// 获取或设置图表对象的宽度
    /// </summary>
    public double Width
    {
        get => _chartObject?.Width ?? 0;
        set
        {
            if (_chartObject != null)
                _chartObject.Width = value;
        }
    }

    /// <summary>
    /// 获取或设置图表对象的高度
    /// </summary>
    public double Height
    {
        get => _chartObject?.Height ?? 0;
        set
        {
            if (_chartObject != null)
                _chartObject.Height = value;
        }
    }

    /// <summary>
    /// 获取或设置图表对象的旋转角度
    /// </summary>
    public double Rotation
    {
        get => _shapeRange?.Rotation ?? 0;
        set
        {
            if (_shapeRange != null)
                _shapeRange.Rotation = (float)value;
        }
    }

    #endregion

    #region 图表属性

    /// <summary>
    /// 图表对象缓存
    /// </summary>
    private IExcelChart _excelChart;

    /// <summary>
    /// 获取图表对象的图表
    /// </summary>
    public IExcelChart Chart => _excelChart ?? (_excelChart = new ExcelChart(_chart));

    /// <summary>
    /// 获取或设置图表对象是否启用宏
    /// </summary>
    public bool EnableMacro
    {
        get => false; // Excel ChartObject不直接支持此属性
        set
        {
            // Excel ChartObject不直接支持此属性
        }
    }

    /// <summary>
    /// 获取图表对象是否为嵌入式图表
    /// </summary>
    public bool IsEmbedded => _chartObject != null;

    /// <summary>
    /// 获取图表对象的图表类型
    /// </summary>
    public int ChartType => _chart != null ? Convert.ToInt32(_chart.ChartType) : 0;

    #endregion

    #region 操作方法

    /// <summary>
    /// 取消保护工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    public void Unprotect(string password = "")
    {
        _chart?.Unprotect(password);
    }

    /// <summary>
    /// 保护工作表
    /// </summary>
    /// <param name="password">保护密码</param>
    /// <param name="drawingObjects">是否保护图形对象</param>
    /// <param name="contents">是否保护内容</param>
    /// <param name="scenarios">是否保护方案</param>
    /// <param name="userInterfaceOnly">是否仅保护用户界面</param>
    public void Protect(string? password = null, bool? drawingObjects = null,
        bool? contents = null, bool? scenarios = null, bool? userInterfaceOnly = null)
    {
        _chart?.Protect(
            password.ComArgsVal(),
            drawingObjects.ComArgsVal(),
            contents.ComArgsVal(),
            scenarios.ComArgsVal(),
            userInterfaceOnly.ComArgsVal());
    }

    public object? OLEObjects(int? index = null)
    {
        if (index != null)
            return _chart?.OLEObjects(index);
        return _chart?.OLEObjects();
    }

    /// <summary>
    /// 选择图表对象
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    public void Select(bool replace = true)
    {
        _chartObject?.Select(replace);
    }

    /// <summary>
    /// 激活图表对象
    /// </summary>
    public void Activate()
    {
        _chartObject?.Activate();
    }

    /// <summary>
    /// 复制图表对象
    /// </summary>
    public void Copy()
    {
        _chartObject?.Copy();
    }

    /// <summary>
    /// 剪切图表对象
    /// </summary>
    public void Cut()
    {
        _chartObject?.Cut();
    }

    /// <summary>
    /// 删除图表对象
    /// </summary>
    public void Delete()
    {
        _chartObject?.Delete();
    }

    /// <summary>
    /// 调整图表对象大小
    /// </summary>
    /// <param name="width">新宽度</param>
    /// <param name="height">新高度</param>
    /// <param name="keepAspectRatio">是否保持纵横比</param>
    public void Resize(double width, double height, bool keepAspectRatio = false)
    {
        if (_chartObject == null) return;

        try
        {
            if (keepAspectRatio)
            {
                double currentRatio = Width / Height;
                double newRatio = width / height;

                if (newRatio > currentRatio)
                {
                    // 以高度为准
                    width = height * currentRatio;
                }
                else
                {
                    // 以宽度为准
                    height = width / currentRatio;
                }
            }

            _chartObject.Width = width;
            _chartObject.Height = height;
        }
        catch
        {
            // 忽略调整大小过程中的异常
        }
    }

    /// <summary>
    /// 移动图表对象
    /// </summary>
    /// <param name="left">新左边距</param>
    /// <param name="top">新顶边距</param>
    public void Move(double left, double top)
    {
        if (_chartObject == null) return;

        try
        {
            _chartObject.Left = left;
            _chartObject.Top = top;
        }
        catch
        {
            // 忽略移动过程中的异常
        }
    }

    /// <summary>
    /// 旋转图表对象
    /// </summary>
    /// <param name="angle">旋转角度</param>
    public void Rotate(double angle)
    {
        if (_shapeRange == null) return;

        try
        {
            _shapeRange.Rotation = (float)angle;
        }
        catch
        {
            // 忽略旋转过程中的异常
        }
    }

    /// <summary>
    /// 将图表对象置于最前面
    /// </summary>
    public void BringToFront()
    {
        _shapeRange?.ZOrder(MsCore.MsoZOrderCmd.msoBringToFront);
    }

    /// <summary>
    /// 将图表对象置于最后面
    /// </summary>
    public void SendToBack()
    {
        _shapeRange?.ZOrder(MsCore.MsoZOrderCmd.msoSendToBack);
    }

    #endregion

    #region 图表操作

    /// <summary>
    /// 设置图表数据源
    /// </summary>
    /// <param name="sourceData">数据源区域</param>
    /// <param name="plotBy">绘制方式</param>
    public void SetSourceData(IExcelRange sourceData, int plotBy = 1)
    {
        if (_chart == null || sourceData == null) return;

        try
        {
            var excelRange = sourceData as ExcelRange;
            if (excelRange?.InternalRange != null)
            {
                _chart.SetSourceData(excelRange.InternalRange, (MsExcel.XlRowCol)plotBy);
            }
        }
        catch
        {
            // 忽略设置数据源过程中的异常
        }
    }

    /// <summary>
    /// 设置图表类型
    /// </summary>
    /// <param name="chartType">图表类型</param>
    public void SetChartType(int chartType)
    {
        if (_chart == null) return;

        try
        {
            _chart.ChartType = (MsExcel.XlChartType)chartType;
        }
        catch
        {
            // 忽略设置图表类型过程中的异常
        }
    }

    /// <summary>
    /// 应用图表布局
    /// </summary>
    /// <param name="layout">布局编号</param>
    public void ApplyLayout(int layout)
    {
        if (_chart == null) return;

        try
        {
            _chart.ApplyLayout(layout);
        }
        catch
        {
            // 忽略应用布局过程中的异常
        }
    }

    /// <summary>
    /// 重新绘制图表
    /// </summary>
    public void Refresh()
    {
        if (_chart == null) return;

        try
        {
            _chart.Refresh();
        }
        catch
        {
            // 忽略重新绘制过程中的异常
        }
    }

    #endregion

    #region 格式设置

    /// <summary>
    /// 设置图表标题
    /// </summary>
    /// <param name="title">标题文本</param>
    public void SetTitle(string title)
    {
        if (_chart == null || string.IsNullOrEmpty(title)) return;

        try
        {
            if (_chart.HasTitle)
            {
                _chart.ChartTitle.Text = title;
            }
            else
            {
                _chart.HasTitle = true;
                _chart.ChartTitle.Text = title;
            }
        }
        catch
        {
            // 忽略设置标题过程中的异常
        }
    }

    /// <summary>
    /// 设置坐标轴标题
    /// </summary>
    /// <param name="axisType">坐标轴类型</param>
    /// <param name="title">标题文本</param>
    public void SetAxisTitle(int axisType, string title)
    {
        if (_chart == null || string.IsNullOrEmpty(title)) return;

        try
        {
            MsExcel.Axis axis = null;
            switch (axisType)
            {
                case 1: // X轴
                    axis = _chart.Axes(MsExcel.XlAxisType.xlCategory) as MsExcel.Axis;
                    break;
                case 2: // Y轴
                    axis = _chart.Axes(MsExcel.XlAxisType.xlValue) as MsExcel.Axis;
                    break;
            }

            if (axis != null)
            {
                if (axis.HasTitle)
                {
                    axis.AxisTitle.Text = title;
                }
                else
                {
                    axis.HasTitle = true;
                    axis.AxisTitle.Text = title;
                }
            }
        }
        catch
        {
            // 忽略设置坐标轴标题过程中的异常
        }
    }

    /// <summary>
    /// 设置图例位置
    /// </summary>
    /// <param name="position">图例位置</param>
    public void SetLegendPosition(int position)
    {
        if (_chart == null) return;

        try
        {
            if (_chart.HasLegend)
            {
                _chart.Legend.Position = (MsExcel.XlLegendPosition)position;
            }
            else
            {
                _chart.HasLegend = true;
                _chart.Legend.Position = (MsExcel.XlLegendPosition)position;
            }
        }
        catch
        {
            // 忽略设置图例位置过程中的异常
        }
    }

    /// <summary>
    /// 设置数据标签
    /// </summary>
    /// <param name="show">是否显示</param>
    public void SetDataLabels(bool show)
    {
        if (_chart == null) return;

        try
        {
            var seriesCollection = _chart.SeriesCollection() as MsExcel.SeriesCollection;
            if (seriesCollection != null)
            {
                for (int i = 1; i <= seriesCollection.Count; i++)
                {
                    try
                    {
                        var series = seriesCollection.Item(i) as MsExcel.Series;
                        if (series != null)
                        {
                            series.HasDataLabels = show;
                        }
                    }
                    catch
                    {
                        // 忽略单个系列设置异常
                    }
                }
            }
        }
        catch
        {
            // 忽略设置数据标签过程中的异常
        }
    }

    /// <summary>
    /// 设置网格线
    /// </summary>
    /// <param name="major">是否显示主要网格线</param>
    /// <param name="minor">是否显示次要网格线</param>
    public void SetGridlines(bool major, bool minor = false)
    {
        if (_chart == null) return;

        try
        {
            // 设置主要网格线
            var valueAxis = _chart.Axes(MsExcel.XlAxisType.xlValue) as MsExcel.Axis;
            if (valueAxis != null)
            {
                valueAxis.HasMajorGridlines = major;
                valueAxis.HasMinorGridlines = minor;
            }

            // 设置次要网格线
            var categoryAxis = _chart.Axes(MsExcel.XlAxisType.xlCategory) as MsExcel.Axis;
            if (categoryAxis != null)
            {
                categoryAxis.HasMajorGridlines = major;
                categoryAxis.HasMinorGridlines = minor;
            }
        }
        catch
        {
            // 忽略设置网格线过程中的异常
        }
    }

    #endregion

    #region 导出和转换



    /// <summary>
    /// 复制图表到新工作表
    /// </summary>
    /// <param name="worksheetName">新工作表名称</param>
    /// <returns>新创建的工作表对象</returns>
    public IExcelWorksheet CopyToNewWorksheet(string worksheetName = "")
    {
        if (_chartObject == null) return null;

        try
        {
            // 获取父工作表
            var parentSheet = _chartObject.Parent as MsExcel.Worksheet;
            if (parentSheet == null)
                return null;

            // 创建新工作表
            var workbook = parentSheet.Parent as MsExcel.Workbook;
            if (workbook == null)
                return null;

            var newSheet = workbook.Worksheets.Add(Type.Missing, parentSheet, Type.Missing, Type.Missing) as MsExcel.Worksheet;

            if (!string.IsNullOrEmpty(worksheetName))
                newSheet.Name = worksheetName;

            // 复制图表到新工作表
            _chartObject.Copy();
            newSheet.Paste();

            return new ExcelWorksheet(newSheet);
        }
        catch
        {
            return null;
        }
    }
    #endregion
}