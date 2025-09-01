//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word.SeriesLines 的封装实现类。
/// </summary>
internal class WordChartSeriesLines : IWordChartSeriesLines
{
    private MsWord.SeriesLines _seriesLines;
    private bool _disposedValue;

    internal WordChartSeriesLines(MsWord.SeriesLines seriesLines)
    {
        _seriesLines = seriesLines ?? throw new ArgumentNullException(nameof(seriesLines));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _seriesLines != null ? new WordApplication(_seriesLines.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _seriesLines?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _seriesLines?.Name ?? string.Empty;
    }

    /// <inheritdoc/>
    public IWordChartFormat Format => _seriesLines?.Format != null ? new WordChartFormat(_seriesLines.Format) : null;

    /// <inheritdoc/>
    public IWordChartBorder Border => _seriesLines?.Border != null ? new WordChartBorder(_seriesLines.Border) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _seriesLines?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _seriesLines?.Delete();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放所有子对象
            (Format as IDisposable)?.Dispose();
            (Border as IDisposable)?.Dispose();

            if (_seriesLines != null)
            {
                Marshal.ReleaseComObject(_seriesLines);
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
