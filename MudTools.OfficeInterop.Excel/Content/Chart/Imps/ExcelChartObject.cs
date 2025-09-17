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

    public IExcelBorder? Border => _chartObject != null ? new ExcelBorder(_chartObject.Border) : null;

    public IExcelInterior? Interior => _chartObject != null ? new ExcelInterior(_chartObject.Interior) : null;

    public IExcelRange? TopLeftCell => _chartObject != null ? new ExcelRange(_chartObject.TopLeftCell) : null;

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

    /// <summary>
    /// 获取图表对象的索引位置
    /// </summary>
    public int Index => _chartObject?.Index ?? 0;


    public bool Visible
    {
        get => _chartObject != null && _chartObject.Visible;
        set
        {
            if (_chartObject != null)
                _chartObject.Visible = value;
        }
    }

    public bool PrintObject
    {
        get => _chartObject.PrintObject;
        set
        {
            if (_chartObject != null)
                _chartObject.PrintObject = value;
        }
    }

    public bool Locked
    {
        get => _chartObject != null && _chartObject.Locked;
        set
        {
            if (_chartObject != null)
                _chartObject.Locked = value;
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
    public IExcelChart Chart => _excelChart ?? (_excelChart = new ExcelChart(_chartObject.Chart));


    /// <summary>
    /// 获取图表对象是否为嵌入式图表
    /// </summary>
    public bool IsEmbedded => _chartObject != null;
    #endregion

    #region 操作方法
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
            if (_chartObject.Parent is not MsExcel.Worksheet parentSheet)
                return null;

            // 创建新工作表
            if (parentSheet.Parent is not MsExcel.Workbook workbook)
                return null;

            var newSheet = workbook.Worksheets.Add(System.Type.Missing, parentSheet, System.Type.Missing, System.Type.Missing) as MsExcel.Worksheet;

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