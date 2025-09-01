//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.ChartTitle 的封装实现类。
/// </summary>
internal class WordChartTitle : IWordChartTitle
{
    private MsWord.ChartTitle _chartTitle;
    private bool _disposedValue;

    internal WordChartTitle(MsWord.ChartTitle chartTitle)
    {
        _chartTitle = chartTitle ?? throw new ArgumentNullException(nameof(chartTitle));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _chartTitle != null ? new WordApplication(_chartTitle.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _chartTitle?.Parent;

    /// <inheritdoc/>
    public string Text
    {
        get => _chartTitle?.Text ?? string.Empty;
        set
        {
            if (_chartTitle != null)
                _chartTitle.Text = value;
        }
    }

    /// <inheritdoc/>
    public XlChartElementPosition Position
    {
        get => _chartTitle?.Position != null ? (XlChartElementPosition)(int)_chartTitle?.Position : XlChartElementPosition.xlChartElementPositionAutomatic;
        set
        {
            if (_chartTitle != null)
                _chartTitle.Position = (MsWord.XlChartElementPosition)(int)value;
        }
    }

    /// <inheritdoc/>
    public string Caption
    {
        get => _chartTitle?.Caption ?? string.Empty;
        set
        {
            if (_chartTitle != null)
                _chartTitle.Caption = value;
        }
    }
    public string Formula
    {
        get => _chartTitle?.Formula ?? string.Empty;
        set
        {
            if (_chartTitle != null)
                _chartTitle.Formula = value;
        }
    }

    public string FormulaLocal
    {
        get => _chartTitle?.FormulaLocal ?? string.Empty;
        set
        {
            if (_chartTitle != null)
                _chartTitle.FormulaLocal = value;
        }
    }

    public string FormulaR1C1
    {
        get => _chartTitle?.FormulaR1C1 ?? string.Empty;
        set
        {
            if (_chartTitle != null)
                _chartTitle.FormulaR1C1 = value;
        }
    }


    /// <inheritdoc/>
    public IWordChartCharacters? Characters => _chartTitle?.Characters != null ? new WordChartCharacters(_chartTitle.Characters) : null;

    /// <inheritdoc/>
    public IWordChartFont? Font => _chartTitle?.Font != null ? new WordChartFont(_chartTitle.Font) : null;

    /// <inheritdoc/>
    public IWordChartFormat? Format => _chartTitle?.Format != null ? new WordChartFormat(_chartTitle.Format) : null;

    /// <inheritdoc/>
    public IWordChartBorder? Border => _chartTitle?.Border != null ? new WordChartBorder(_chartTitle.Border) : null;


    /// <inheritdoc/>
    public XlHAlign HorizontalAlignment
    {
        get => _chartTitle?.HorizontalAlignment != null ? (XlHAlign)(int)_chartTitle?.HorizontalAlignment : XlHAlign.xlHAlignLeft;
        set
        {
            if (_chartTitle != null)
                _chartTitle.HorizontalAlignment = (MsWord.XlHAlign)(int)value;
        }
    }

    /// <inheritdoc/>
    public XlVAlign VerticalAlignment
    {
        get => _chartTitle?.VerticalAlignment != null ? (XlVAlign)(int)_chartTitle?.VerticalAlignment : XlVAlign.xlVAlignJustify;
        set
        {
            if (_chartTitle != null)
                _chartTitle.VerticalAlignment = (MsWord.XlVAlign)(int)value;
        }
    }

    /// <inheritdoc/>
    public bool AutoScaleFont
    {
        get => _chartTitle.AutoScaleFont.ConvertToBool();
        set
        {
            if (_chartTitle != null)
                _chartTitle.AutoScaleFont = value;
        }
    }

    /// <inheritdoc/>
    public bool IncludeInLayout
    {
        get => _chartTitle.IncludeInLayout;
        set
        {
            if (_chartTitle != null)
                _chartTitle.IncludeInLayout = value;
        }
    }

    public int ReadingOrder
    {
        get => _chartTitle?.ReadingOrder ?? 0;
        set
        {
            if (_chartTitle != null)
                _chartTitle.ReadingOrder = value;
        }
    }


    /// <inheritdoc/>
    public double Left
    {
        get => _chartTitle?.Left ?? 0.0;
        set
        {
            if (_chartTitle != null)
                _chartTitle.Left = value;
        }
    }

    /// <inheritdoc/>
    public double Top
    {
        get => _chartTitle?.Top ?? 0.0;
        set
        {
            if (_chartTitle != null)
                _chartTitle.Top = value;
        }
    }

    /// <inheritdoc/>
    public double Width
    {
        get => _chartTitle?.Width ?? 0.0;
    }

    /// <inheritdoc/>
    public double Height
    {
        get => _chartTitle?.Height ?? 0.0;
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _chartTitle?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _chartTitle?.Delete();
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
            (Format as IDisposable)?.Dispose();

            if (_chartTitle != null)
            {
                Marshal.ReleaseComObject(_chartTitle);
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