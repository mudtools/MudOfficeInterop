//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Floor 的封装实现类。
/// </summary>
internal class WordFloor : IWordFloor
{
    private MsWord.Floor _floor;
    private bool _disposedValue;

    internal WordFloor(MsWord.Floor floor)
    {
        _floor = floor ?? throw new ArgumentNullException(nameof(floor));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _floor != null ? new WordApplication(_floor.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _floor?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _floor?.Name ?? string.Empty;
    }

    /// <inheritdoc/>
    public int Thickness
    {
        get => _floor?.Thickness ?? 0;
        set
        {
            if (_floor != null)
                _floor.Thickness = value;
        }
    }

    #endregion

    #region 对象属性实现  
    /// <inheritdoc/>
    public IWordInterior? Interior => _floor?.Interior != null ? new WordInterior(_floor.Interior) : null;

    /// <inheritdoc/>
    public IWordChartFillFormat? Fill => _floor?.Fill != null ? new WordChartFillFormat(_floor.Fill) : null;

    /// <inheritdoc/>
    public IWordChartBorder? Border => _floor?.Border != null ? new WordChartBorder(_floor.Border) : null;

    /// <inheritdoc/>
    public IWordChartFormat? Format => _floor?.Format != null ? new WordChartFormat(_floor.Format) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _floor?.Select();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {
            if (_floor != null)
            {
                Marshal.ReleaseComObject(_floor);
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
