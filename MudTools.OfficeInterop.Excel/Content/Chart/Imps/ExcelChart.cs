//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// Excel Chart 对象的二次封装实现类
/// 实现 IExcelChart 接口，提供对 Microsoft.Office.Interop.Excel.Chart 的安全访问和操作
/// </summary>
internal partial class ExcelChart : IExcelChart
{
    #region 私有字段

    /// <summary>
    /// 内部持有的 Microsoft.Office.Interop.Excel.Chart 对象引用
    /// </summary>
    internal MsExcel.Chart? _chart;

    /// <summary>
    /// 标记对象是否已被释放，用于防止重复释放
    /// </summary>
    private bool _disposedValue = false;

    #endregion

    #region 构造函数

    /// <summary>
    /// 初始化 ExcelChart 实例
    /// </summary>
    /// <param name="chart">要封装的 Microsoft.Office.Interop.Excel.Chart 对象</param>
    /// <exception cref="ArgumentNullException">当 chart 为 null 时抛出</exception>
    internal ExcelChart(MsExcel.Chart chart)
    {
        _chart = chart ?? throw new ArgumentNullException(nameof(chart));
        _chartEvents_Event = chart;
        InitializeEvents();
    }

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取或设置图表的名称
    /// </summary>
    public string Name
    {
        get => _chart != null ? _chart.Name : string.Empty;
        set
        {
            if (_chart != null)
                _chart.Name = value;
        }
    }

    public XlSheetType Type => _chart != null ? _chart.Type.ObjectConvertEnum<XlSheetType>(XlSheetType.xlChart) : XlSheetType.xlChart;

    /// <summary>
    /// 获取图表在集合中的索引位置 (从1开始)
    /// </summary>
    public int Index => _chart != null ? _chart.Index : 0;

    /// <summary>
    /// 获取或设置图表的类型 (使用 XlChartType 枚举对应的 int 值)
    /// </summary>
    public MsoChartType ChartType
    {
        get => _chart != null ? _chart.ChartType.EnumConvert(MsoChartType.xl3DColumn) : MsoChartType.xl3DColumn;
        set
        {
            if (_chart != null)
                _chart.ChartType = value.EnumConvert(MsExcel.XlChartType.xl3DColumn);
        }
    }

    /// <summary>
    /// 获取或设置图表是否可见
    /// </summary>
    public XlSheetVisibility Visible
    {
        get => _chart != null ? _chart.Visible.EnumConvert(XlSheetVisibility.xlSheetVisible) : XlSheetVisibility.xlSheetVisible;
        set
        {
            if (_chart != null)
                _chart.Visible = value.EnumConvert(MsExcel.XlSheetVisibility.xlSheetVisible);
        }
    }

    public bool IsVisible
    {
        get => _chart != null && _chart.Visible == MsExcel.XlSheetVisibility.xlSheetVisible;
        set
        {
            if (_chart != null)
                _chart.Visible = value ? (MsExcel.XlSheetVisibility.xlSheetVisible) : (MsExcel.XlSheetVisibility.xlSheetHidden);
        }
    }

    /// <summary>
    /// 获取图表内容是否被保护
    /// </summary>
    public bool IsProtected => _chart != null && _chart.ProtectionMode;

    public bool ProtectContents => _chart != null && _chart.ProtectionMode;

    public bool ProtectionMode => _chart != null && _chart.ProtectionMode;


    /// <summary>
    /// 获取图表的父对象
    /// </summary>
    public object? Parent
    {
        get
        {
            if (_chart?.Parent == null)
            {
                return null;
            }
            if (_chart.Parent is MsExcel.ChartObject chatObj)
            {
                return new ExcelChartObject(chatObj);
            }
            if (_chart.Parent is MsExcel.Worksheet worksheet)
            {
                return new ExcelWorksheet(worksheet);
            }
            if (_chart.Parent is MsExcel.Workbook workbook)
            {
                return new ExcelWorkbook(workbook);
            }
            return _chart.Parent;
        }
    }

    public IExcelWorkbook? ParentWorkbook
    {
        get
        {
            if (_chart?.Parent == null)
            {
                return null;
            }
            if (_chart.Parent is MsExcel.Workbook workbook)
            {
                return new ExcelWorkbook(workbook);
            }
            return null;
        }
    }

    public string? ParentName
    {
        get
        {
            if (_chart?.Parent == null)
            {
                return null;
            }
            if (_chart.Parent is MsExcel.ChartObject chatObj)
            {
                return chatObj.Name;
            }
            if (_chart.Parent is MsExcel.Worksheet worksheet)
            {
                return worksheet.Name;
            }
            if (_chart.Parent is MsExcel.Workbook workbook)
            {
                return workbook.Name;
            }
            return null;
        }
    }

    /// <summary>
    /// 获取图表所在的 Excel Application 对象
    /// </summary>
    public IExcelApplication? Application => _chart != null ? new ExcelApplication(_chart.Application) : null;

    /// <summary>
    /// 获取图表的代码名称 (只读)
    /// </summary>
    public string CodeName => _chart != null ? _chart.CodeName : "";

    #endregion

    #region 位置和大小 (IExcelChart)  

    /// <summary>
    /// 获取或设置图表的旋转角度
    /// </summary>
    public double Rotation
    {
        get => _chart != null ? Convert.ToDouble(_chart.Rotation) : 0;
        set
        {
            if (_chart != null)
                _chart.Rotation = Convert.ToInt32(value);
        }
    }

    #endregion

    #region 数据源 (IExcelChart)
    /// <summary>
    /// 获取或设置数据绘制方式 (行/列优先)
    /// </summary>
    public XlRowCol PlotBy
    {
        get => _chart != null ? (XlRowCol)_chart.PlotBy.EnumConvert(XlRowCol.xlColumns) : XlRowCol.xlColumns;
        set
        {
            if (_chart != null)
                _chart.PlotBy = value.EnumConvert(MsExcel.XlRowCol.xlColumns);
        }
    }

    /// <summary>
    /// 获取或设置图表是否有标题
    /// </summary>
    public bool HasTitle
    {
        get => _chart != null && _chart.HasTitle;
        set
        {
            if (_chart != null)
                _chart.HasTitle = value;
        }
    }

    /// <summary>
    /// 获取或设置图表标题文本
    /// </summary>
    public string ChartTitle
    {
        get => _chart != null ? _chart.ChartTitle.Text : "";
        set
        {
            if (_chart != null)
            {
                HasTitle = true;
                _chart.ChartTitle.Text = value;
            }
        }
    }

    /// <summary>
    /// 获取或设置图表是否有图例
    /// </summary>
    public bool HasLegend
    {
        get => _chart != null && _chart.HasLegend;
        set
        {
            if (_chart != null)
                _chart.HasLegend = value;
        }
    }

    /// <summary>
    /// 获取或设置图例的位置 (使用 XlLegendPosition 枚举对应的 int 值)
    /// </summary>
    public XlLegendPosition LegendPosition
    {
        get => _chart != null ? _chart.Legend.Position.EnumConvert(XlLegendPosition.xlLegendPositionRight) : XlLegendPosition.xlLegendPositionRight;
        set
        {
            if (_chart != null)
                _chart.Legend.Position = value.EnumConvert(MsExcel.XlLegendPosition.xlLegendPositionRight);
        }
    }

    #endregion

    #region 图表元素 (IExcelChart)

    public IExcelShapes? Shapes => _chart != null ? new ExcelShapes(_chart.Shapes) : null;

    /// <summary>
    /// 获取图表的绘图区对象
    /// </summary>
    public IExcelPlotArea? PlotArea => _chart != null ? new ExcelPlotArea(_chart.PlotArea) : null;

    /// <summary>
    /// 获取图表的图表区对象
    /// </summary>
    public IExcelChartArea? ChartArea => _chart != null ? new ExcelChartArea(_chart.ChartArea) : null;


    public IExcelAxes? Axes(XlAxisType? axisType = null, XlAxisGroup axisGroup = XlAxisGroup.xlPrimary)
    {
        if (_chart == null)
            return null;
        var charAxesObj = _chart.Axes(axisType.ComArgsConvert(x => x.EnumConvert(MsExcel.XlAxisType.xlValue)),
                          axisGroup.EnumConvert(MsExcel.XlAxisGroup.xlPrimary));
        if (charAxesObj != null && charAxesObj is MsExcel.Axes charAxes)
            return new ExcelAxes(charAxes);
        return null;
    }

    /// <summary>
    /// 获取图表的图表标题对象
    /// </summary>
    public IExcelChartTitle? ChartTitleObject => _chart != null && HasTitle ? new ExcelChartTitle(_chart.ChartTitle) : null;

    /// <summary>
    /// 获取图表的图例对象
    /// </summary>
    public IExcelLegend? Legend => _chart != null && HasLegend ? new ExcelLegend(_chart.Legend) : null;

    /// <summary>
    /// 获取图表的数据标签集合 (通常在 Series 上)
    /// </summary>
    public IExcelDataTable? DataTable => _chart != null ? new ExcelDataTable(_chart.DataTable) : null;

    /// <summary>
    /// 页面设置对象缓存
    /// </summary>
    private IExcelPageSetup? _pageSetup;

    /// <summary>
    /// 获取工作表的页面设置对象
    /// </summary>
    public IExcelPageSetup? PageSetup
    {
        get
        {
            if (_chart == null)
            {
                return null;
            }
            _pageSetup ??= new ExcelPageSetup(_chart.PageSetup);
            return _pageSetup;
        }
    }

    /// <summary>
    /// 超链接集合缓存
    /// </summary>
    private IExcelHyperlinks? _hyperlinks;

    /// <summary>
    /// 获取工作表的超链接集合
    /// </summary>
    public IExcelHyperlinks? Hyperlinks
    {
        get
        {
            if (_chart == null)
            {
                return null;
            }
            _hyperlinks ??= new ExcelHyperlinks(_chart.Hyperlinks);
            return _hyperlinks;
        }
    }

    #endregion

    #region 图表设置 (IExcelChart)

    /// <summary>
    /// 获取或设置是否在图表中显示数据表
    /// </summary>
    public bool HasDataTable
    {
        get => _chart != null && _chart.HasDataTable;
        set
        {
            if (_chart != null)
                _chart.HasDataTable = value;
        }
    }


    /// <summary>
    /// 获取或设置图表样式编号
    /// </summary>
    public XlChartType ChartStyle
    {
        get => _chart != null ? _chart.ChartStyle.ObjectConvertEnum(XlChartType.xl3DColumn) : XlChartType.xl3DColumn;
        set
        {
            if (_chart != null)
                _chart.ChartStyle = value.EnumConvert(MsExcel.XlChartType.xl3DColumn);
        }
    }
    #endregion

    #region 操作方法 (IExcelChart)

    /// <summary>
    /// 复制工作表
    /// </summary>
    /// <param name="before">复制到指定工作表之前</param>
    /// <param name="after">复制到指定工作表之后</param>
    public void Copy(IExcelComSheet? before = null, IExcelComSheet? after = null)
    {
        if (_chart == null) return;

        _chart.Copy(
            before is ExcelChart beforeSheet ? beforeSheet._chart : System.Type.Missing,
            after is ExcelChart afterSheet ? afterSheet._chart : System.Type.Missing
        );
    }

    /// <summary>
    /// 将工作表另存为xlsx文件。
    /// </summary>
    /// <param name="filePath"></param>
    public void SaveAs(string filePath)
    {
        _chart?.SaveAs(filePath);
    }

    /// <summary>
    /// 选择图表
    /// </summary>
    /// <param name="replace">是否替换当前选择</param>
    public void Select(bool replace = true)
    {
        _chart?.Select(replace);
    }

    /// <summary>
    /// 激活图表
    /// </summary>
    public void Activate()
    {
        _chart?.Activate();
    }

    /// <summary>
    /// 复制图表
    /// </summary>
    public void Copy()
    {
        _chart?.Copy();
    }


    /// <summary>
    /// 删除图表
    /// </summary>
    public void Delete()
    {
        _chart?.Delete();
    }


    /// <summary>
    /// 旋转图表
    /// </summary>
    /// <param name="angle">旋转角度</param>
    public void Rotate(double angle)
    {
        Rotation = angle;
    }

    /// <summary>
    /// 移动工作表
    /// </summary>
    /// <param name="before">移动到指定工作表之前</param>
    /// <param name="after">移动到指定工作表之后</param>
    public void Move(IExcelComSheet? before = null, IExcelComSheet? after = null)
    {
        if (_chart == null) return;

        _chart.Move(
            before is ExcelChart beforeSheet ? beforeSheet._chart : System.Type.Missing,
            after is ExcelChart afterSheet ? afterSheet._chart : System.Type.Missing
        );
    }
    #endregion

    #region 图表操作 (IExcelChart)

    public void SetBackgroundPicture(string filename)
    {
        _chart?.SetBackgroundPicture(filename);
    }

    /// <summary>
    /// 设置图表的数据源
    /// </summary>
    /// <param name="sourceData">数据源区域</param>
    /// <param name="plotBy">绘制方式 (1=列, 2=行)</param>
    public void SetSourceData(IExcelRange sourceData, XlRowCol plotBy = XlRowCol.xlRows)
    {
        if (sourceData is ExcelRange excelRange)
        {
            _chart?.SetSourceData(excelRange.InternalRange, (MsExcel.XlRowCol)(int)plotBy);
        }
        else
        {
            throw new ArgumentException("SourceData must be of type ExcelRange or compatible.");
        }
    }

    /// <summary>
    /// 应用指定的图表布局
    /// </summary>
    /// <param name="layout">布局编号</param>
    public void ApplyLayout(int layout)
    {
        _chart?.ApplyLayout(layout);
    }

    /// <summary>
    /// 刷新图表数据和显示
    /// </summary>
    public void Refresh()
    {
        _chart?.Refresh();
    }

    /// <summary>
    /// 清除所有图表内容
    /// </summary>
    public void ClearAll()
    {
        _chart?.ChartArea.ClearFormats();
        _chart?.ChartArea.ClearContents();
        _chart?.ChartArea.Clear();
    }

    /// <summary>
    /// 清除图表内容
    /// </summary>
    public void Clear()
    {
        _chart?.ChartArea.Clear();
    }
    public void ClearContents()
    {
        _chart?.ChartArea.ClearContents();
    }

    /// <summary>
    /// 清除图表格式
    /// </summary>
    public void ClearFormats()
    {
        _chart?.ChartArea.ClearFormats();
    }

    #endregion

    #region 格式设置 (IExcelChart)

    /// <summary>
    /// 获取图表的数据系列集合
    /// </summary>
    public IExcelSeriesCollection? SeriesCollection()
    {
        if (_chart?.SeriesCollection() is MsExcel.SeriesCollection series)
            return new ExcelSeriesCollection(series);
        return null;
    }

    /// <summary>
    /// 获取图表的数据系列集合
    /// </summary>
    public IExcelSeries? SeriesCollection(int index)
    {
        if (_chart?.SeriesCollection(index) is MsExcel.Series series)
            return new ExcelSeries(series);
        return null;
    }

    /// <summary>
    /// 设置图表标题
    /// </summary>
    /// <param name="title">标题文本</param>
    public void SetTitle(string title)
    {
        HasTitle = true;
        if (_chart != null)
            _chart.ChartTitle.Text = title;
    }


    /// <summary>
    /// 设置图例位置
    /// </summary>
    /// <param name="position">图例位置 (使用 XlLegendPosition 枚举对应的 int 值)</param>
    public void SetLegendPosition(XlLegendPosition position)
    {
        HasLegend = true;
        if (_chart != null)
            _chart.Legend.Position = (MsExcel.XlLegendPosition)(int)position;
    }

    /// <summary>
    /// 设置数据标签显示
    /// </summary>
    /// <param name="show">是否显示</param>
    public void SetDataLabels(bool show)
    {
        if (show && _chart != null)
        {
            if (_chart.SeriesCollection() is not MsExcel.SeriesCollection seriesColl)
                return;
            for (int i = 1; i <= seriesColl.Count; i++)
            {
                MsExcel.Series series = seriesColl.Item(i);
                series.ApplyDataLabels(Type: MsExcel.XlDataLabelsType.xlDataLabelsShowValue);
            }
        }
    }

    /// <summary>
    /// 设置图表背景色 (图表区)
    /// </summary>
    /// <param name="color">RGB 颜色值</param>
    public void SetBackgroundColor(int color)
    {
        if (_chart != null)
            _chart.ChartArea.Format.Fill.ForeColor.RGB = color;
    }

    /// <summary>
    /// 设置图表前景色 (图表区边框)
    /// </summary>
    /// <param name="color">RGB 颜色值</param>
    public void SetForegroundColor(int color)
    {
        if (_chart != null)
            _chart.ChartArea.Format.Line.ForeColor.RGB = color;
    }

    #endregion

    #region 导出和转换 (IExcelChart)

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
        if (index != null && _chart != null)
            return _chart.OLEObjects(index);
        return _chart?.OLEObjects();
    }

    /// <summary>
    /// 将图表导出为图片文件
    /// </summary>
    /// <param name="filename">导出文件路径</param>
    /// <param name="format">图片格式 (如 "PNG", "JPG")</param>
    /// <param name="overwrite">是否覆盖已存在文件</param>
    /// <returns>导出是否成功</returns>
    public bool ExportToImage(string filename, string format = "png", bool overwrite = true)
    {
        try
        {
            if (System.IO.File.Exists(filename) && !overwrite)
            {
                return false;
            }
            _chart?.Export(filename, format, false);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 获取图表图片的字节数据
    /// </summary>
    /// <param name="format">图片格式 (如 "PNG", "JPG")</param>
    /// <returns>图片字节数组</returns>
    public byte[] GetImageBytes(string format = "png")
    {
        // 通过临时文件方式获取字节
        string tempPath = System.IO.Path.GetTempFileName() + "." + format;
        try
        {
            _chart?.Export(tempPath, format, false);
            if (System.IO.File.Exists(tempPath))
            {
                byte[] bytes = System.IO.File.ReadAllBytes(tempPath);
                System.IO.File.Delete(tempPath);
                return bytes;
            }
        }
        catch
        {
            if (System.IO.File.Exists(tempPath))
            {
                try { System.IO.File.Delete(tempPath); } catch { }
            }
        }
        return []; // Return empty array on failure
    }
    #endregion

    #region 高级功能 (IExcelChart)
    public void PrintPreview()
    {
        _chart?.PrintPreview();
    }

    /// <summary>
    /// 打印图表
    /// </summary>
    /// <param name="preview">是否进行打印预览</param>
    public void PrintOut(bool preview = false)
    {
        if (preview)
        {
            _chart?.PrintPreview();
        }
        else
        {
            _chart?.PrintOutEx();
        }
    }

    #endregion

    #region IDisposable Support

    /// <summary>
    /// 释放资源的核心方法
    /// </summary>
    /// <param name="disposing">如果为 true，则同时释放托管和非托管资源；如果为 false，则仅释放非托管资源。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            _pageSetup?.Dispose();
            _hyperlinks?.Dispose();
            DisconnectEvents();
            if (_chart != null)
                Marshal.ReleaseComObject(_chart);
            _pageSetup = null;
            _chart = null;
            _hyperlinks = null;
            _chartEvents_Event = null;
        }
        _disposedValue = true;
    }

    /// <summary>
    /// 终结器 (析构函数)，防止资源未被释放
    /// </summary>
    ~ExcelChart()
    {
        Dispose(disposing: false);
    }

    /// <summary>
    /// 公开的 Dispose 方法，实现 IDisposable 接口
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    #endregion
}
