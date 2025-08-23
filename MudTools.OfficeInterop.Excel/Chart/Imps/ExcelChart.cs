//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;


/// <summary>
/// Excel Chart 对象的二次封装实现类
/// 实现 IExcelChart 接口，提供对 Microsoft.Office.Interop.Excel.Chart 的安全访问和操作
/// </summary>
internal class ExcelChart : IExcelChart
{
    #region 私有字段

    /// <summary>
    /// 内部持有的 Microsoft.Office.Interop.Excel.Chart 对象引用
    /// </summary>
    internal MsExcel.Chart _chart;

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
    }

    #endregion

    #region 基础属性

    /// <summary>
    /// 获取或设置图表的名称
    /// </summary>
    public string Name
    {
        get => _chart.Name;
        set => _chart.Name = value;
    }

    /// <summary>
    /// 获取图表在集合中的索引位置 (从1开始)
    /// </summary>
    public int Index => _chart.Index;

    /// <summary>
    /// 获取或设置图表的类型 (使用 XlChartType 枚举对应的 int 值)
    /// </summary>
    public MsoChartType ChartType
    {
        get => (MsoChartType)_chart.ChartType;
        set => _chart.ChartType = (MsExcel.XlChartType)value;
    }

    /// <summary>
    /// 获取或设置图表是否可见
    /// </summary>
    public XlSheetVisibility Visible
    {
        get => (XlSheetVisibility)_chart.Visible;
        set => _chart.Visible = (MsExcel.XlSheetVisibility)value;
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
    public bool IsProtected => _chart.ProtectContents;

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
    public IExcelApplication Application => new ExcelApplication(_chart.Application);

    /// <summary>
    /// 获取图表的代码名称 (只读)
    /// </summary>
    public string CodeName => _chart.CodeName;

    #endregion

    #region 位置和大小 (IExcelChart)  

    /// <summary>
    /// 获取或设置图表的旋转角度
    /// </summary>
    public double Rotation
    {
        get => Convert.ToDouble(_chart.Rotation);
        set => _chart.Rotation = value;
    }

    #endregion

    #region 数据源 (IExcelChart)

    /// <summary>
    /// 获取或设置图表的数据源区域
    /// 注意：获取 SourceData 可能需要解析图表公式，此处为简化实现。
    /// </summary>
    public IExcelRange SourceData
    {
        get
        {
            // 简化处理：返回 null 或需要解析 _chart.Formula 或其他属性
            // 实际应用中可能需要更复杂的逻辑来获取原始 Range
            return null;
        }
        set
        {
            if (value is ExcelRange excelRange)
            {
                _chart.SetSourceData(excelRange.InternalRange);
            }
            else
            {
                throw new ArgumentException("SourceData must be of type ExcelRange or compatible.");
            }
        }
    }

    /// <summary>
    /// 获取或设置数据绘制方式 (行/列优先)
    /// </summary>
    public int PlotBy
    {
        get => (int)_chart.PlotBy;
        set => _chart.PlotBy = (MsExcel.XlRowCol)value;
    }

    /// <summary>
    /// 获取或设置图表是否有标题
    /// </summary>
    public bool HasTitle
    {
        get => _chart.HasTitle;
        set => _chart.HasTitle = value;
    }

    /// <summary>
    /// 获取或设置图表标题文本
    /// </summary>
    public string ChartTitle
    {
        get => HasTitle ? _chart.ChartTitle.Text : null;
        set
        {
            HasTitle = true;
            _chart.ChartTitle.Text = value;
        }
    }

    /// <summary>
    /// 获取或设置图表是否有图例
    /// </summary>
    public bool HasLegend
    {
        get => _chart.HasLegend;
        set => _chart.HasLegend = value;
    }

    /// <summary>
    /// 获取或设置图例的位置 (使用 XlLegendPosition 枚举对应的 int 值)
    /// </summary>
    public int LegendPosition
    {
        get => HasLegend ? (int)_chart.Legend.Position : -1; // -1 表示无图例
        set
        {
            HasLegend = true;
            _chart.Legend.Position = (MsExcel.XlLegendPosition)value;
        }
    }

    #endregion

    #region 图表元素 (IExcelChart)

    public IExcelShapes Shapes => new ExcelShapes(_chart.Shapes);

    /// <summary>
    /// 获取图表的绘图区对象
    /// </summary>
    public IExcelPlotArea PlotArea => new ExcelPlotArea(_chart.PlotArea);

    /// <summary>
    /// 获取图表的图表区对象
    /// </summary>
    public IExcelChartArea ChartArea => new ExcelChartArea(_chart.ChartArea);

    /// <summary>
    /// 获取图表的坐标轴集合
    /// </summary>
    public IExcelAxes Axes => new ExcelAxes((MsExcel.Axes)_chart.Axes());

    /// <summary>
    /// 获取图表的图表标题对象
    /// </summary>
    public IExcelChartTitle ChartTitleObject => HasTitle ? new ExcelChartTitle(_chart.ChartTitle) : null;

    /// <summary>
    /// 获取图表的图例对象
    /// </summary>
    public IExcelLegend Legend => HasLegend ? new ExcelLegend(_chart.Legend) : null;

    /// <summary>
    /// 获取图表的数据标签集合 (通常在 Series 上)
    /// </summary>
    public IExcelDataTable DataTable => new ExcelDataTable(_chart.DataTable);

    #endregion

    #region 图表设置 (IExcelChart)

    /// <summary>
    /// 获取或设置是否在图表中显示数据表
    /// </summary>
    public bool HasDataTable
    {
        get => _chart.HasDataTable;
        set => _chart.HasDataTable = value;
    }


    /// <summary>
    /// 获取或设置图表样式编号
    /// </summary>
    public int ChartStyle
    {
        get => (int)_chart.ChartStyle;
        set => _chart.ChartStyle = value;
    }
    #endregion

    #region 操作方法 (IExcelChart)

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
        _chart.Select(replace);
    }

    /// <summary>
    /// 激活图表
    /// </summary>
    public void Activate()
    {
        _chart.Activate();
    }

    /// <summary>
    /// 复制图表
    /// </summary>
    public void Copy()
    {
        _chart.Copy();
    }

    /// <summary>
    /// 删除图表
    /// </summary>
    public void Delete()
    {
        _chart.Delete();
    }


    /// <summary>
    /// 旋转图表
    /// </summary>
    /// <param name="angle">旋转角度</param>
    public void Rotate(double angle)
    {
        Rotation = angle;
    }
    #endregion

    #region 图表操作 (IExcelChart)

    /// <summary>
    /// 设置图表的数据源
    /// </summary>
    /// <param name="sourceData">数据源区域</param>
    /// <param name="plotBy">绘制方式 (1=列, 2=行)</param>
    public void SetSourceData(IExcelRange sourceData, int plotBy = 1)
    {
        if (sourceData is ExcelRange excelRange)
        {
            _chart.SetSourceData(excelRange.InternalRange, (MsExcel.XlRowCol)plotBy);
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
        _chart.ApplyLayout(layout);
    }

    /// <summary>
    /// 刷新图表数据和显示
    /// </summary>
    public void Refresh()
    {
        _chart.Refresh();
    }

    /// <summary>
    /// 清除图表内容
    /// </summary>
    public void Clear()
    {
        _chart.ChartArea.Clear();
    }

    /// <summary>
    /// 清除图表格式
    /// </summary>
    public void ClearFormats()
    {
        _chart.ChartArea.ClearFormats();
    }

    #endregion

    #region 格式设置 (IExcelChart)

    /// <summary>
    /// 获取图表的数据系列集合
    /// </summary>
    public IExcelSeriesCollection? SeriesCollection()
    {
        var series = _chart.SeriesCollection() as MsExcel.SeriesCollection;
        if (series != null)
            return new ExcelSeriesCollection(series);
        return null;
    }

    /// <summary>
    /// 获取图表的数据系列集合
    /// </summary>
    public IExcelSeries? SeriesCollection(int index)
    {
        var series = _chart.SeriesCollection(index) as MsExcel.Series;
        if (series != null)
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
        _chart.ChartTitle.Text = title;
    }


    /// <summary>
    /// 设置图例位置
    /// </summary>
    /// <param name="position">图例位置 (使用 XlLegendPosition 枚举对应的 int 值)</param>
    public void SetLegendPosition(XlLegendPosition position)
    {
        HasLegend = true;
        _chart.Legend.Position = (MsExcel.XlLegendPosition)position;
    }

    /// <summary>
    /// 设置数据标签显示
    /// </summary>
    /// <param name="show">是否显示</param>
    public void SetDataLabels(bool show)
    {
        if (show)
        {
            MsExcel.SeriesCollection? seriesColl = _chart.SeriesCollection() as MsExcel.SeriesCollection;
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
        _chart.ChartArea.Format.Fill.ForeColor.RGB = color;
    }

    /// <summary>
    /// 设置图表前景色 (图表区边框)
    /// </summary>
    /// <param name="color">RGB 颜色值</param>
    public void SetForegroundColor(int color)
    {
        _chart.ChartArea.Format.Line.ForeColor.RGB = color;
    }

    #endregion

    #region 导出和转换 (IExcelChart)

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
            _chart.Export(filename, format, false); // interactive = false
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
            _chart.Export(tempPath, format, false);
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

    /// <summary>
    /// 打印图表
    /// </summary>
    /// <param name="preview">是否进行打印预览</param>
    public void PrintOut(bool preview = false)
    {
        if (preview)
        {
            _chart.PrintPreview();
        }
        else
        {
            _chart.PrintOutEx();
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
            try
            {
                // 释放底层COM对象
                if (_chart != null)
                    Marshal.ReleaseComObject(_chart);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _chart = null;
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
