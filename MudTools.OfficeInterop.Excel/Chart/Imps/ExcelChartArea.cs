//
// 懒人Excel工具箱 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;
/// <summary>
/// Excel ChartArea 对象的二次封装实现类
/// 实现 IExcelChartArea 接口
/// </summary>
internal class ExcelChartArea : IExcelChartArea
{
    private MsExcel.ChartArea _chartArea;
    private bool _disposedValue = false;

    internal ExcelChartArea(MsExcel.ChartArea chartArea)
    {
        _chartArea = chartArea ?? throw new ArgumentNullException(nameof(chartArea));
    }

    #region 基础属性
    public string Name
    {
        get => _chartArea.Name;
    }


    public object Parent => _chartArea.Parent;

    public IExcelApplication Application => new ExcelApplication(_chartArea.Application);
    #endregion

    #region 位置和大小
    public double Left
    {
        get => (double)_chartArea.Left;
        set => _chartArea.Left = value;
    }

    public double Top
    {
        get => (double)_chartArea.Top;
        set => _chartArea.Top = value;
    }

    public double Width
    {
        get => (double)_chartArea.Width;
        set => _chartArea.Width = value;
    }

    public double Height
    {
        get => (double)_chartArea.Height;
        set => _chartArea.Height = value;
    }
    #endregion

    #region 格式设置
    public IExcelFont Font => new ExcelFont(_chartArea.Font);

    public bool AutoScaleFont
    {
        get => Convert.ToBoolean(_chartArea.AutoScaleFont);
        set => _chartArea.AutoScaleFont = value;
    }

    public IExcelFillFormat Fill => new ExcelFillFormat(_chartArea.Format.Fill);
    public IExcelBorder Border => new ExcelBorder(_chartArea.Border);

    public bool Shadow
    {
        get => _chartArea.Shadow;
        set => _chartArea.Shadow = value;
    }
    #endregion

    #region 操作方法  

    public void Clear()
    {
        _chartArea.Clear();
    }

    public void ClearFormats()
    {
        _chartArea.ClearFormats();
    }

    public void ClearAll()
    {
        _chartArea.Clear();
        _chartArea.ClearFormats();
    }

    public void Copy()
    {
        _chartArea.Copy();
    }

    public BoundingBox GetBoundingBox()
    {
        return new BoundingBox
        {
            Left = Left,
            Top = Top,
            Width = Width,
            Height = Height
        };
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                // 释放托管状态(托管对象)
            }

            if (_chartArea != null)
            {
                try
                {
                    while (Marshal.ReleaseComObject(_chartArea) > 0) { }
                }
                catch { }
                _chartArea = null;
            }

            _disposedValue = true;
        }
    }

    ~ExcelChartArea()
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
