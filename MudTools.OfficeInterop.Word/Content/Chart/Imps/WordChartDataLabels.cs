//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.DataLabels 的封装实现类。
/// </summary>
internal class WordChartDataLabels : IWordChartDataLabels
{
    private MsWord.DataLabels _dataLabels;
    private bool _disposedValue;

    internal WordChartDataLabels(MsWord.DataLabels dataLabels)
    {
        _dataLabels = dataLabels ?? throw new ArgumentNullException(nameof(dataLabels));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _dataLabels != null ? new WordApplication(_dataLabels.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _dataLabels?.Parent;

    /// <inheritdoc/>
    public int Count => _dataLabels?.Count ?? 0;

    /// <inheritdoc/>
    public bool AutoText
    {
        get => _dataLabels?.AutoText ?? true;
        set
        {
            if (_dataLabels != null)
                _dataLabels.AutoText = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowLegendKey
    {
        get => _dataLabels?.ShowLegendKey ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ShowLegendKey = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowValue
    {
        get => _dataLabels?.ShowValue ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ShowValue = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowCategoryName
    {
        get => _dataLabels?.ShowCategoryName ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ShowCategoryName = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowSeriesName
    {
        get => _dataLabels?.ShowSeriesName ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ShowSeriesName = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowPercentage
    {
        get => _dataLabels?.ShowPercentage ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ShowPercentage = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowBubbleSize
    {
        get => _dataLabels?.ShowBubbleSize ?? false;
        set
        {
            if (_dataLabels != null)
                _dataLabels.ShowBubbleSize = value;
        }
    }


    /// <inheritdoc/>
    public XlDataLabelPosition Position
    {
        get => _dataLabels?.Position != null ? (XlDataLabelPosition)(int)_dataLabels?.Position : XlDataLabelPosition.xlLabelPositionLeft;
        set
        {
            if (_dataLabels != null) _dataLabels.Position = (MsWord.XlDataLabelPosition)(int)value;
        }
    }

    /// <inheritdoc/>
    public XlHAlign HorizontalAlignment
    {
        get => _dataLabels?.HorizontalAlignment != null ? (XlHAlign)(int)_dataLabels?.HorizontalAlignment : XlHAlign.xlHAlignLeft;
        set
        {
            if (_dataLabels != null) _dataLabels.HorizontalAlignment = (MsWord.XlHAlign)(int)value;
        }
    }

    /// <inheritdoc/>
    public XlVAlign VerticalAlignment
    {
        get => _dataLabels?.VerticalAlignment != null ? (XlVAlign)(int)_dataLabels?.VerticalAlignment : XlVAlign.xlVAlignJustify;
        set
        {
            if (_dataLabels != null) _dataLabels.VerticalAlignment = (MsWord.XlVAlign)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool AutoScaleFont
    {
        get => _dataLabels.AutoScaleFont.ConvertToBool();
        set
        {
            if (_dataLabels != null)
                _dataLabels.AutoScaleFont = value;
        }
    }

    #endregion

    #region 对象属性实现

    /// <inheritdoc/>
    public IWordChartFont Font => _dataLabels?.Font != null ? new WordChartFont(_dataLabels.Font) : null;

    /// <inheritdoc/>
    public IWordInterior Interior => _dataLabels?.Interior != null ? new WordInterior(_dataLabels.Interior) : null;

    /// <inheritdoc/>
    public IWordChartFillFormat Fill => _dataLabels?.Fill != null ? new WordChartFillFormat(_dataLabels.Fill) : null;

    /// <inheritdoc/>
    public IWordChartBorder Border => _dataLabels?.Border != null ? new WordChartBorder(_dataLabels.Border) : null;

    /// <inheritdoc/>
    public IWordChartFormat Format => _dataLabels?.Format != null ? new WordChartFormat(_dataLabels.Format) : null;

    #endregion

    #region 索引器实现

    /// <inheritdoc/>
    public IWordDataLabel this[int index]
    {
        get
        {
            if (index < 1 || index > Count) return null;

            try
            {
                var comDataLabel = _dataLabels[index];
                return new WordDataLabel(comDataLabel);
            }
            catch
            {
                return null;
            }
        }
    }

    #endregion

    #region 方法实现
    /// <inheritdoc/>
    public void Delete()
    {
        _dataLabels?.Delete();
    }

    /// <inheritdoc/>
    public void Select()
    {
        _dataLabels?.Select();
    }

    /// <inheritdoc/>
    public List<int> GetIndexes()
    {
        var indexes = new List<int>();
        for (int i = 1; i <= Count; i++)
        {
            indexes.Add(i);
        }
        return indexes;
    }

    #endregion

    #region 枚举支持

    /// <inheritdoc/>
    public IEnumerator<IWordDataLabel> GetEnumerator()
    {
        for (int i = 1; i <= Count; i++)
        {
            yield return this[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放所有子对象
            (Font as IDisposable)?.Dispose();
            (Interior as IDisposable)?.Dispose();
            (Fill as IDisposable)?.Dispose();
            (Border as IDisposable)?.Dispose();
            (Format as IDisposable)?.Dispose();

            if (_dataLabels != null)
            {
                Marshal.ReleaseComObject(_dataLabels);
            }
        }

        _disposedValue = true;
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}