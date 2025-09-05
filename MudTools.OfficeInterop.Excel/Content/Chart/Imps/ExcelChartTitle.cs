//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel ChartTitle 对象的二次封装实现类
/// 实现 IExcelChartTitle 接口
/// </summary>
internal class ExcelChartTitle : IExcelChartTitle
{
    private MsExcel.ChartTitle _chartTitle;
    private bool _disposedValue = false;

    internal ExcelChartTitle(MsExcel.ChartTitle chartTitle)
    {
        _chartTitle = chartTitle ?? throw new ArgumentNullException(nameof(chartTitle));
    }

    #region 基础属性
    public string Name
    {
        get => _chartTitle.Name;
    }


    public string Text
    {
        get => _chartTitle.Text;
        set => _chartTitle.Text = value;
    }

    public object Parent => _chartTitle.Parent;

    public IExcelApplication Application => new ExcelApplication(_chartTitle.Application);
    #endregion

    #region 位置和大小
    public double Left
    {
        get => _chartTitle.Left;
        set => _chartTitle.Left = value;
    }

    public double Top
    {
        get => _chartTitle.Top;
        set => _chartTitle.Top = value;
    }

    public double Width
    {
        get => _chartTitle.Width;
    }

    public double Height
    {
        get => _chartTitle.Height;
    }
    #endregion

    #region 格式设置
    public IExcelFont Font => new ExcelFont(_chartTitle.Font);

    public IExcelChartFormat Format => new ExcelChartFormat(_chartTitle.Format);

    /// <summary>
    /// 获取样式的内部格式对象
    /// </summary>
    public IExcelInterior Interior => new ExcelInterior(_chartTitle.Interior);


    public bool AutoScaleFont
    {
        get => Convert.ToBoolean(_chartTitle.AutoScaleFont);
        set => _chartTitle.AutoScaleFont = value;
    }

    public IExcelChartFillFormat Fill => new ExcelChartFillFormat(_chartTitle.Fill);

    public int HorizontalAlignment
    {
        get => (int)_chartTitle.HorizontalAlignment;
        set => _chartTitle.HorizontalAlignment = (MsExcel.XlHAlign)value;
    }

    public int VerticalAlignment
    {
        get => (int)_chartTitle.VerticalAlignment;
        set => _chartTitle.VerticalAlignment = (MsExcel.XlVAlign)value;
    }

    public int ReadingOrder
    {
        get => _chartTitle.ReadingOrder;
        set => _chartTitle.ReadingOrder = value;
    }

    public int Orientation
    {
        get => (int)_chartTitle.Orientation;
        set => _chartTitle.Orientation = value; // value 应为 MsExcel.XlOrientation 枚举值对应的 int
    }
    #endregion

    #region 操作方法
    public void Select()
    {
        _chartTitle.Select();
    }

    public void Delete()
    {
        _chartTitle.Delete();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            try
            {
                // 释放底层COM对象
                if (_chartTitle != null)
                    Marshal.ReleaseComObject(_chartTitle);
            }
            catch
            {
                // 忽略释放过程中的异常
            }
            _chartTitle = null;
        }
        _disposedValue = true;
    }

    ~ExcelChartTitle()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
    #endregion
}
