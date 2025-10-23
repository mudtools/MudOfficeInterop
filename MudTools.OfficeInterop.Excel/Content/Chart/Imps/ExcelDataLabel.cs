namespace MudTools.OfficeInterop.Excel.Imps;

/// <summary>
/// <see cref="IExcelDataLabel"/> 接口的内部实现类。
/// 负责包装 Microsoft.Office.Interop.Excel.DataLabel COM 对象，并管理其生命周期。
/// </summary>
internal class ExcelDataLabel : IExcelDataLabel
{
    internal MsExcel.DataLabel? _dataLabel;
    private static readonly ILog log = LogManager.GetLogger(typeof(ExcelDataLabel));
    private bool _disposedValue = false;

    /// <summary>
    /// 初始化 <see cref="ExcelDataLabel"/> 类的新实例。
    /// </summary>
    /// <param name="dataLabel">要包装的原始 COM 对象。</param>
    /// <exception cref="ArgumentNullException">当 <paramref name="dataLabel"/> 为 null 时抛出。</exception>
    internal ExcelDataLabel(MsExcel.DataLabel dataLabel)
    {
        _dataLabel = dataLabel ?? throw new ArgumentNullException(nameof(dataLabel));
    }

    public object? Parent => _dataLabel?.Parent;
    public IExcelApplication? Application => _dataLabel != null ? new ExcelApplication(_dataLabel.Application) : null;

    public string Text
    {
        get => _dataLabel?.Text ?? string.Empty;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Text = value;
        }
    }

    public bool AutoText
    {
        get => _dataLabel?.AutoText ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.AutoText = value;
        }
    }

    public bool ShowLegendKey
    {
        get => _dataLabel?.ShowLegendKey ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ShowLegendKey = value;
        }
    }

    public bool ShowPercentage
    {
        get => _dataLabel?.ShowPercentage ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ShowPercentage = value;
        }
    }

    public bool ShowSeriesName
    {
        get => _dataLabel?.ShowSeriesName ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ShowSeriesName = value;
        }
    }

    public bool ShowCategoryName
    {
        get => _dataLabel?.ShowCategoryName ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ShowCategoryName = value;
        }
    }

    public bool ShowValue
    {
        get => _dataLabel?.ShowValue ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ShowValue = value;
        }
    }

    public XlDataLabelPosition Position
    {
        get => _dataLabel?.Position.EnumConvert(XlDataLabelPosition.xlLabelPositionBestFit) ?? XlDataLabelPosition.xlLabelPositionBestFit;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Position = value.EnumConvert(MsExcel.XlDataLabelPosition.xlLabelPositionBestFit);
        }
    }

    public string NumberFormat
    {
        get => _dataLabel?.NumberFormat?.ToString() ?? string.Empty;
        set
        {
            if (_dataLabel != null)
                _dataLabel.NumberFormat = value;
        }
    }

    public bool NumberFormatLinked
    {
        get => _dataLabel?.NumberFormatLinked ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.NumberFormatLinked = value;
        }
    }

    public IExcelChartFormat? Format => _dataLabel != null ? new ExcelChartFormat(_dataLabel.Format) : null;
    public IExcelBorder? Border => _dataLabel != null ? new ExcelBorder(_dataLabel.Border) : null;
    public IExcelInterior? Interior => _dataLabel != null ? new ExcelInterior(_dataLabel.Interior) : null;
    public IExcelFont? Font => _dataLabel != null ? new ExcelFont(_dataLabel.Font) : null;

    public object? HorizontalAlignment
    {
        get => _dataLabel?.HorizontalAlignment;
        set
        {
            if (_dataLabel != null)
                _dataLabel.HorizontalAlignment = value;
        }
    }

    public object? VerticalAlignment
    {
        get => _dataLabel?.VerticalAlignment;
        set
        {
            if (_dataLabel != null)
                _dataLabel.VerticalAlignment = value;
        }
    }

    public int ReadingOrder
    {
        get => _dataLabel?.ReadingOrder ?? 0;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ReadingOrder = value;
        }
    }

    public object? Orientation
    {
        get => _dataLabel?.Orientation;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Orientation = value;
        }
    }

    public double Left
    {
        get => _dataLabel?.Left ?? 0.0;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Left = value;
        }
    }

    public double Top
    {
        get => _dataLabel?.Top ?? 0.0;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Top = value;
        }
    }

    public double Width
    {
        get => _dataLabel?.Width ?? 0.0;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Width = value;
        }
    }

    public double Height
    {
        get => _dataLabel?.Height ?? 0.0;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Height = value;
        }
    }

    /// <summary>
    /// 删除该数据标签。
    /// </summary>
    /// <exception cref="ObjectDisposedException">当对象已被释放时抛出。</exception>
    public void Delete()
    {
        if (_dataLabel == null) throw new ObjectDisposedException(nameof(ExcelDataLabel));
        try
        {
            _dataLabel.Delete();
        }
        catch (Exception ex)
        {
            log.Error($"删除数据标签失败: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// 选择该数据标签。
    /// </summary>
    /// <exception cref="ObjectDisposedException">当对象已被释放时抛出。</exception>
    public void Select()
    {
        if (_dataLabel == null) throw new ObjectDisposedException(nameof(ExcelDataLabel));
        try
        {
            _dataLabel.Select();
        }
        catch (Exception ex)
        {
            log.Error($"选择数据标签失败: {ex.Message}");
            throw;
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _dataLabel != null)
        {
            Marshal.ReleaseComObject(_dataLabel);
            _dataLabel = null;
        }

        _disposedValue = true;
    }

    ~ExcelDataLabel()
    {
        Dispose(disposing: false);
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}