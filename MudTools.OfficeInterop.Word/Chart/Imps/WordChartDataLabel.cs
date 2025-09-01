//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.DataLabel 的封装实现类。
/// </summary>
internal class WordChartDataLabel : IWordChartDataLabel
{
    private MsWord.DataLabel _dataLabel;
    private bool _disposedValue;

    internal WordChartDataLabel(MsWord.DataLabel dataLabel)
    {
        _dataLabel = dataLabel ?? throw new ArgumentNullException(nameof(dataLabel));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _dataLabel != null ? new WordApplication(_dataLabel.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _dataLabel?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _dataLabel?.Name ?? string.Empty;
    }


    /// <inheritdoc/>
    public string Text
    {
        get => _dataLabel?.Text ?? string.Empty;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Text = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowLegendKey
    {
        get => _dataLabel?.ShowLegendKey ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ShowLegendKey = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowValue
    {
        get => _dataLabel?.ShowValue ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ShowValue = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowCategoryName
    {
        get => _dataLabel?.ShowCategoryName ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ShowCategoryName = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowSeriesName
    {
        get => _dataLabel?.ShowSeriesName ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ShowSeriesName = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowPercentage
    {
        get => _dataLabel?.ShowPercentage ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ShowPercentage = value;
        }
    }

    /// <inheritdoc/>
    public bool ShowBubbleSize
    {
        get => _dataLabel?.ShowBubbleSize ?? false;
        set
        {
            if (_dataLabel != null)
                _dataLabel.ShowBubbleSize = value;
        }
    }

    /// <inheritdoc/>
    public bool AutoText
    {
        get => _dataLabel?.AutoText ?? true;
        set
        {
            if (_dataLabel != null)
                _dataLabel.AutoText = value;
        }
    }

    /// <inheritdoc/>
    public XlDataLabelPosition Position
    {
        get => _dataLabel?.Position != null ? (XlDataLabelPosition)(int)_dataLabel?.Position : XlDataLabelPosition.xlLabelPositionLeft;
        set
        {
            if (_dataLabel != null) _dataLabel.Position = (MsWord.XlDataLabelPosition)(int)value;
        }
    }

    /// <inheritdoc/>
    public XlHAlign HorizontalAlignment
    {
        get => _dataLabel?.HorizontalAlignment != null ? (XlHAlign)(int)_dataLabel?.HorizontalAlignment : XlHAlign.xlHAlignLeft;
        set
        {
            if (_dataLabel != null) _dataLabel.HorizontalAlignment = (MsWord.XlHAlign)(int)value;
        }
    }

    /// <inheritdoc/>
    public XlVAlign VerticalAlignment
    {
        get => _dataLabel?.VerticalAlignment != null ? (XlVAlign)(int)_dataLabel?.VerticalAlignment : XlVAlign.xlVAlignJustify;
        set
        {
            if (_dataLabel != null) _dataLabel.VerticalAlignment = (MsWord.XlVAlign)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool AutoScaleFont
    {
        get => _dataLabel.AutoScaleFont.ConvertToBool();
        set
        {
            if (_dataLabel != null)
                _dataLabel.AutoScaleFont = value;
        }
    }

    /// <inheritdoc/>
    public double Left
    {
        get => _dataLabel?.Left ?? 0.0;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Left = value;
        }
    }

    /// <inheritdoc/>
    public double Top
    {
        get => _dataLabel?.Top ?? 0.0;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Top = value;
        }
    }

    /// <inheritdoc/>
    public double Width
    {
        get => _dataLabel?.Width ?? 0.0;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Width = value;
        }
    }

    /// <inheritdoc/>
    public double Height
    {
        get => _dataLabel?.Height ?? 0.0;
        set
        {
            if (_dataLabel != null)
                _dataLabel.Height = value;
        }
    }

    #endregion

    #region 对象属性实现

    /// <inheritdoc/>
    public IWordChartCharacters? Characters => _dataLabel?.Characters != null ? new WordChartCharacters(_dataLabel.Characters) : null;

    /// <inheritdoc/>
    public IWordChartFont? Font => _dataLabel?.Font != null ? new WordChartFont(_dataLabel.Font) : null;

    /// <inheritdoc/>
    public IWordInterior? Interior => _dataLabel?.Interior != null ? new WordInterior(_dataLabel.Interior) : null;

    /// <inheritdoc/>
    public IWordChartFillFormat? Fill => _dataLabel?.Fill != null ? new WordChartFillFormat(_dataLabel.Fill) : null;

    /// <inheritdoc/>
    public IWordChartBorder? Border => _dataLabel?.Border != null ? new WordChartBorder(_dataLabel.Border) : null;

    /// <inheritdoc/>
    public IWordChartFormat? Format => _dataLabel?.Format != null ? new WordChartFormat(_dataLabel.Format) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _dataLabel?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _dataLabel?.Delete();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放所有子对象
            (Characters as IDisposable)?.Dispose();
            (Font as IDisposable)?.Dispose();
            (Interior as IDisposable)?.Dispose();
            (Fill as IDisposable)?.Dispose();
            (Border as IDisposable)?.Dispose();
            (Format as IDisposable)?.Dispose();

            if (_dataLabel != null)
            {
                Marshal.ReleaseComObject(_dataLabel);
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