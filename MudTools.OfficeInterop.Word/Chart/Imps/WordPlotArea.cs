//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.PlotArea 的封装实现类。
/// </summary>
internal class WordPlotArea : IWordPlotArea
{
    private MsWord.PlotArea _plotArea;
    private bool _disposedValue;

    internal WordPlotArea(MsWord.PlotArea plotArea)
    {
        _plotArea = plotArea ?? throw new ArgumentNullException(nameof(plotArea));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _plotArea != null ? new WordApplication(_plotArea.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _plotArea?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _plotArea?.Name ?? string.Empty;
    }

    /// <inheritdoc/>
    public double Left
    {
        get => _plotArea?.Left ?? 0.0;
        set
        {
            if (_plotArea != null)
                _plotArea.Left = value;
        }
    }

    /// <inheritdoc/>
    public double Top
    {
        get => _plotArea?.Top ?? 0.0;
        set
        {
            if (_plotArea != null)
                _plotArea.Top = value;
        }
    }

    /// <inheritdoc/>
    public double Width
    {
        get => _plotArea?.Width ?? 0.0;
        set
        {
            if (_plotArea != null)
                _plotArea.Width = value;
        }
    }

    /// <inheritdoc/>
    public double Height
    {
        get => _plotArea?.Height ?? 0.0;
        set
        {
            if (_plotArea != null)
                _plotArea.Height = value;
        }
    }

    /// <inheritdoc/>
    public object InteriorColor
    {
        get => _plotArea?.Interior.Color;
        set
        {
            if (_plotArea?.Interior != null)
                _plotArea.Interior.Color = value;
        }
    }

    /// <inheritdoc/>
    public object BorderColor
    {
        get => _plotArea?.Border.Color;
        set
        {
            if (_plotArea?.Border != null)
                _plotArea.Border.Color = value;
        }
    }

    /// <inheritdoc/>
    public bool Fill
    {
        get => _plotArea?.Fill.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_plotArea?.Fill != null)
                _plotArea.Fill.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }
    #endregion

    #region 对象属性实现
    /// <inheritdoc/>
    public IWordInterior Interior => _plotArea?.Interior != null ? new WordInterior(_plotArea.Interior) : null;

    /// <inheritdoc/>
    public IWordChartFillFormat FillFormat => _plotArea?.Fill != null ? new WordChartFillFormat(_plotArea.Fill) : null;

    /// <inheritdoc/>
    public IWordChartBorder Border => _plotArea?.Border != null ? new WordChartBorder(_plotArea.Border) : null;

    /// <inheritdoc/>
    public IWordChartFormat Format => _plotArea?.Format != null ? new WordChartFormat(_plotArea.Format) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _plotArea?.Select();
    }

    /// <inheritdoc/>
    public void ClearFormats()
    {
        _plotArea?.ClearFormats();
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_plotArea != null)
            {
                Marshal.ReleaseComObject(_plotArea);
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