//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.HiLoLines 的封装实现类。
/// </summary>
internal class WordChartHiLoLines : IWordChartHiLoLines
{
    private MsWord.HiLoLines _hiLoLines;
    private bool _disposedValue;

    internal WordChartHiLoLines(MsWord.HiLoLines hiLoLines)
    {
        _hiLoLines = hiLoLines ?? throw new ArgumentNullException(nameof(hiLoLines));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _hiLoLines != null ? new WordApplication(_hiLoLines.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _hiLoLines?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _hiLoLines?.Name ?? string.Empty;
    }


    /// <inheritdoc/>
    public IWordChartFormat Format => _hiLoLines?.Format != null ? new WordChartFormat(_hiLoLines.Format) : null;

    /// <inheritdoc/>
    public IWordChartBorder Border => _hiLoLines?.Border != null ? new WordChartBorder(_hiLoLines.Border) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _hiLoLines?.Select();
    }

    /// <inheritdoc/>
    public void Delete()
    {
        _hiLoLines?.Delete();
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

            if (_hiLoLines != null)
            {
                Marshal.ReleaseComObject(_hiLoLines);
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