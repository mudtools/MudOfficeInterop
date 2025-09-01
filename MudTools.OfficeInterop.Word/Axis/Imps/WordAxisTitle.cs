//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.AxisTitle 的封装实现类。
/// </summary>
internal class WordAxisTitle : IWordAxisTitle
{
    private MsWord.AxisTitle _axisTitle;
    private bool _disposedValue;

    internal WordAxisTitle(MsWord.AxisTitle axisTitle)
    {
        _axisTitle = axisTitle ?? throw new ArgumentNullException(nameof(axisTitle));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _axisTitle != null ? new WordApplication(_axisTitle.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _axisTitle?.Parent;

    /// <inheritdoc/>
    public string Text
    {
        get => _axisTitle?.Text ?? string.Empty;
        set
        {
            if (_axisTitle != null)
                _axisTitle.Text = value;
        }
    }

    /// <inheritdoc/>
    public string Name => _axisTitle?.Name ?? string.Empty;

    /// <inheritdoc/>
    public bool AutoScaleFont
    {
        get => _axisTitle.AutoScaleFont.ConvertToBool();
        set
        {
            if (_axisTitle != null)
                _axisTitle.AutoScaleFont = value;
        }
    }

    /// <inheritdoc/>
    public double Left
    {
        get => _axisTitle?.Left ?? 0.0;
        set
        {
            if (_axisTitle != null)
                _axisTitle.Left = value;
        }
    }

    /// <inheritdoc/>
    public double Top
    {
        get => _axisTitle?.Top ?? 0.0;
        set
        {
            if (_axisTitle != null)
                _axisTitle.Top = value;
        }
    }

    /// <inheritdoc/>
    public double Width
    {
        get => _axisTitle?.Width ?? 0.0;
    }

    /// <inheritdoc/>
    public double Height
    {
        get => _axisTitle?.Height ?? 0.0;
    }

    /// <inheritdoc/>
    public XlHAlign HorizontalAlignment
    {
        get => _axisTitle?.HorizontalAlignment != null ? (XlHAlign)(int)_axisTitle?.HorizontalAlignment : XlHAlign.xlHAlignLeft;
        set
        {
            if (_axisTitle != null)
                _axisTitle.HorizontalAlignment = (MsWord.XlHAlign)(int)value;
        }
    }

    /// <inheritdoc/>
    public XlVAlign VerticalAlignment
    {
        get => _axisTitle?.VerticalAlignment != null ? (XlVAlign)(int)_axisTitle?.VerticalAlignment : XlVAlign.xlVAlignJustify;
        set
        {
            if (_axisTitle != null)
                _axisTitle.VerticalAlignment = (MsWord.XlVAlign)(int)value;
        }
    }

    #endregion

    #region 对象属性实现

    /// <inheritdoc/>
    public IWordChartCharacters Characters => _axisTitle?.Characters != null ? new WordChartCharacters(_axisTitle.Characters) : null;

    /// <inheritdoc/>
    public IWordChartFont Font => _axisTitle?.Font != null ? new WordChartFont(_axisTitle.Font) : null;

    /// <inheritdoc/>
    public IWordInterior Interior => _axisTitle?.Interior != null ? new WordInterior(_axisTitle.Interior) : null;

    /// <inheritdoc/>
    public IWordChartFillFormat Fill => _axisTitle?.Fill != null ? new WordChartFillFormat(_axisTitle.Fill) : null;

    /// <inheritdoc/>
    public IWordChartBorder Border => _axisTitle?.Border != null ? new WordChartBorder(_axisTitle.Border) : null;

    /// <inheritdoc/>
    public IWordChartFormat Format => _axisTitle?.Format != null ? new WordChartFormat(_axisTitle.Format) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _axisTitle?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _axisTitle?.Delete();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_axisTitle != null)
            {
                Marshal.ReleaseComObject(_axisTitle);
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