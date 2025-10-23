//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// <see cref="IExcelDataLabels"/> 接口的内部实现类。
/// 负责包装 Microsoft.Office.Interop.Excel.DataLabels COM 对象，并管理其生命周期及子对象的生命周期。
/// </summary>
internal class ExcelDataLabels : IExcelDataLabels
{
    internal MsExcel.DataLabels? _dataLabels;
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelDataLabel));
    private DisposableList _disposables = new();
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 <see cref="ExcelDataLabels"/> 类的新实例。
    /// </summary>
    /// <param name="dataLabels">要包装的原始 COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 <paramref name="dataLabels"/> 为 null 时抛出。</exception>
    internal ExcelDataLabels(MsExcel.DataLabels dataLabels)
    {
        _dataLabels = dataLabels ?? throw new ArgumentNullException(nameof(dataLabels));
    }

    public int Count => _dataLabels?.Count ?? 0;

    public IExcelDataLabel? this[int index]
    {
        get
        {
            if (_dataLabels == null) return null;
            var label = new ExcelDataLabel(_dataLabels.Item(index));
            _disposables.Add(label);
            return label;
        }
    }

    public object? Parent => _dataLabels?.Parent;
    public IExcelApplication? Application => _dataLabels != null ? new ExcelApplication(_dataLabels.Application) : null;

    public bool ShowLegendKey
    {
        get => _dataLabels?.ShowLegendKey ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ShowLegendKey = value;
        }
    }

    public bool ShowPercentage
    {
        get => _dataLabels?.ShowPercentage ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ShowPercentage = value;
        }
    }

    public bool ShowSeriesName
    {
        get => _dataLabels?.ShowSeriesName ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ShowSeriesName = value;
        }
    }

    public bool ShowCategoryName
    {
        get => _dataLabels?.ShowCategoryName ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ShowCategoryName = value;
        }
    }

    public bool ShowValue
    {
        get => _dataLabels?.ShowValue ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ShowValue = value;
        }
    }

    public XlDataLabelPosition Position
    {
        get => _dataLabels?.Position.EnumConvert(XlDataLabelPosition.xlLabelPositionBestFit) ?? XlDataLabelPosition.xlLabelPositionBestFit;
        set
        {
            if (_dataLabels != null)
                _dataLabels.Position = value.EnumConvert(MsExcel.XlDataLabelPosition.xlLabelPositionBestFit);
        }
    }

    public string NumberFormat
    {
        get => _dataLabels?.NumberFormat?.ToString() ?? string.Empty;
        set
        {
            if (_dataLabels != null)
                _dataLabels.NumberFormat = value;
        }
    }

    public bool NumberFormatLinked
    {
        get => _dataLabels?.NumberFormatLinked ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.NumberFormatLinked = value;
        }
    }

    public IExcelChartFormat? Format => _dataLabels != null ? new ExcelChartFormat(_dataLabels.Format) : null;
    public IExcelBorder? Border => _dataLabels != null ? new ExcelBorder(_dataLabels.Border) : null;
    public IExcelInterior? Interior => _dataLabels != null ? new ExcelInterior(_dataLabels.Interior) : null;
    public IExcelFont? Font => _dataLabels != null ? new ExcelFont(_dataLabels.Font) : null;

    public object? HorizontalAlignment
    {
        get => _dataLabels?.HorizontalAlignment;
        set
        {
            if (_dataLabels != null)
                _dataLabels.HorizontalAlignment = value;
        }
    }

    public object? VerticalAlignment
    {
        get => _dataLabels?.VerticalAlignment;
        set
        {
            if (_dataLabels != null)
                _dataLabels.VerticalAlignment = value;
        }
    }

    public int ReadingOrder
    {
        get => _dataLabels?.ReadingOrder ?? 0;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ReadingOrder = value;
        }
    }

    public object? Orientation
    {
        get => _dataLabels?.Orientation;
        set
        {
            if (_dataLabels != null)
                _dataLabels.Orientation = value;
        }
    }

    /// <summary>
    /// 删除集合中的所有数据标签。
    /// </summary>
    /// <exception cref="ObjectDisposedException">当对象已被释放时抛出。</exception>
    public void Delete()
    {
        if (_dataLabels == null) throw new ObjectDisposedException(nameof(ExcelDataLabels));
        try
        {
            _dataLabels.Delete();
        }
        catch (Exception ex)
        {
            log.Error($"删除数据标签集合失败: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// 将单个数据标签的内容和格式应用到系列中的所有其他数据标签。
    /// </summary>
    /// <param name="index">要传播的单个数据标签的索引。</param>
    /// <exception cref="ObjectDisposedException">当对象已被释放时抛出。</exception>
    public void Propagate(int index)
    {
        if (_dataLabels == null) throw new ObjectDisposedException(nameof(ExcelDataLabels));
        try
        {
            _dataLabels.Propagate(index);
        }
        catch (Exception ex)
        {
            log.Error($"传播数据标签格式失败: {ex.Message}");
            throw;
        }
    }

    public IEnumerator<IExcelDataLabel> GetEnumerator()
    {
        if (_dataLabels == null)
            yield break;

        for (int i = 1; i <= _dataLabels.Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _dataLabels != null)
        {
            Marshal.ReleaseComObject(_dataLabels);
            _dataLabels = null;
            _disposables.Dispose();
        }

        _disposedValue = true;
    }

    ~ExcelDataLabels()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}