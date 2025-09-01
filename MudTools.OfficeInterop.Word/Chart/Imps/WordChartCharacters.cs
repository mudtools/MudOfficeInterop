//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.ChartCharacters 的封装实现类。
/// </summary>
internal class WordChartCharacters : IWordChartCharacters
{
    private MsWord.ChartCharacters _chartCharacters;
    private bool _disposedValue;

    internal WordChartCharacters(MsWord.ChartCharacters chartCharacters)
    {
        _chartCharacters = chartCharacters ?? throw new ArgumentNullException(nameof(chartCharacters));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _chartCharacters != null ? new WordApplication(_chartCharacters.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _chartCharacters?.Parent;

    /// <inheritdoc/>
    public int Count => _chartCharacters?.Count ?? 0;

    /// <inheritdoc/>
    public IWordChartFont Font => _chartCharacters?.Font != null ? new WordChartFont(_chartCharacters.Font) : null;

    /// <inheritdoc/>
    public string Text
    {
        get => _chartCharacters?.Text ?? string.Empty;
        set
        {
            if (_chartCharacters != null)
                _chartCharacters.Text = value;
        }
    }

    /// <inheritdoc/>
    public string Caption
    {
        get => _chartCharacters?.Caption ?? string.Empty;
        set
        {
            if (_chartCharacters != null)
                _chartCharacters.Caption = value;
        }
    }

    /// <inheritdoc/>
    public string PhoneticCharacters
    {
        get => _chartCharacters?.PhoneticCharacters ?? string.Empty;
        set
        {
            if (_chartCharacters != null)
                _chartCharacters.PhoneticCharacters = value;
        }
    }
    #endregion

    #region 方法实现
    /// <inheritdoc/>
    public void Delete()
    {
        _chartCharacters?.Delete();
    }

    /// <inheritdoc/>
    public void Insert(string text)
    {
        _chartCharacters?.Insert(text);
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            // 释放字体对象
            (Font as IDisposable)?.Dispose();

            if (_chartCharacters != null)
            {
                Marshal.ReleaseComObject(_chartCharacters);
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