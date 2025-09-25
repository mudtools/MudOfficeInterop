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
    private MsExcel.ChartArea? _chartArea;
    private bool _disposedValue = false;

    internal ExcelChartArea(MsExcel.ChartArea chartArea)
    {
        _chartArea = chartArea ?? throw new ArgumentNullException(nameof(chartArea));
    }

    #region 基础属性
    public string Name
    {
        get => _chartArea != null ? _chartArea.Name : string.Empty;
    }


    public object? Parent => _chartArea != null ? _chartArea.Parent : null;

    public IExcelApplication? Application => _chartArea != null ? new ExcelApplication(_chartArea.Application) : null;
    #endregion

    #region 位置和大小
    public double Left
    {
        get => _chartArea != null ? _chartArea.Left : 0;
        set
        {
            if (_chartArea != null)
                _chartArea.Left = value;
        }
    }

    public double Top
    {
        get => _chartArea != null ? _chartArea.Top : 0;
        set
        {
            if (_chartArea != null)
                _chartArea.Top = value;
        }
    }

    public double Width
    {
        get => _chartArea != null ? _chartArea.Width : 0;
        set
        {
            if (_chartArea != null)
                _chartArea.Width = value;
        }
    }

    public double Height
    {
        get => _chartArea != null ? _chartArea.Height : 0;
        set
        {
            if (_chartArea != null)
                _chartArea.Height = value;
        }
    }
    #endregion

    #region 格式设置
    public IExcelFont? Font
    {
        get
        {
            if (_chartArea == null)
                return null;
            return new ExcelFont(_chartArea.Font);
        }
    }

    public bool AutoScaleFont
    {
        get => _chartArea != null ? _chartArea.AutoScaleFont.ConvertToBool() : false;
        set
        {
            if (_chartArea != null)
                _chartArea.AutoScaleFont = value;
        }
    }

    public IExcelChartFormat Format
    {
        get
        {
            if (_chartArea == null)
                return null;
            return new ExcelChartFormat(_chartArea.Format);
        }
    }

    public IExcelChartFillFormat? Fill
    {
        get
        {
            if (_chartArea == null)
                return null;
            return new ExcelChartFillFormat(_chartArea.Fill);
        }
    }
    public IExcelBorder? Border
    {
        get
        {
            if (_chartArea == null)
                return null;
            return new ExcelBorder(_chartArea.Border);
        }
    }

    public IExcelInterior? Interior
    {
        get
        {
            if (_chartArea == null)
                return null;
            return new ExcelInterior(_chartArea.Interior);
        }
    }

    public bool RoundedCorners
    {
        get => _chartArea != null ? _chartArea.RoundedCorners : false;
        set
        {
            if (_chartArea != null)
                _chartArea.RoundedCorners = value;
        }
    }

    public bool Shadow
    {
        get => _chartArea != null ? _chartArea.Shadow : false;
        set
        {
            if (_chartArea != null)
                _chartArea.Shadow = value;
        }
    }
    #endregion

    #region 操作方法  

    public void Clear()
    {
        _chartArea?.Clear();
    }

    public void ClearFormats()
    {
        _chartArea?.ClearFormats();
    }

    public void ClearAll()
    {
        _chartArea?.Clear();
        _chartArea?.ClearFormats();
    }

    public void Copy()
    {
        _chartArea?.Copy();
    }
    #endregion

    #region IDisposable Support
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                if (_chartArea != null)
                {
                    Marshal.ReleaseComObject(_chartArea);
                    _chartArea = null;
                }
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