//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.Walls 的封装实现类。
/// </summary>
internal class WordWalls : IWordWalls
{
    private MsWord.Walls _walls;
    private bool _disposedValue;

    internal WordWalls(MsWord.Walls walls)
    {
        _walls = walls ?? throw new ArgumentNullException(nameof(walls));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _walls != null ? new WordApplication(_walls.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object Parent => _walls?.Parent;

    /// <inheritdoc/>
    public string Name
    {
        get => _walls?.Name ?? string.Empty;
    }


    /// <inheritdoc/>
    public int Thickness
    {
        get => _walls?.Thickness ?? 0;
        set
        {
            if (_walls != null)
                _walls.Thickness = value;
        }
    }

    /// <inheritdoc/>
    public bool PictureType
    {
        get => (MsWord.XlPictureAppearance)_walls?.PictureType == MsWord.XlPictureAppearance.xlScreen;
        set
        {
            if (_walls != null)
                _walls.PictureType = value ? MsWord.XlPictureAppearance.xlScreen : MsWord.XlPictureAppearance.xlPrinter;
        }
    }

    #endregion

    #region 对象属性实现   
    /// <inheritdoc/>
    public IWordInterior? Interior => _walls?.Interior != null ? new WordInterior(_walls.Interior) : null;

    /// <inheritdoc/>
    public IWordChartFillFormat? Fill => _walls?.Fill != null ? new WordChartFillFormat(_walls.Fill) : null;

    /// <inheritdoc/>
    public IWordChartBorder? Border => _walls?.Border != null ? new WordChartBorder(_walls.Border) : null;

    /// <inheritdoc/>
    public IWordChartFormat? Format => _walls?.Format != null ? new WordChartFormat(_walls.Format) : null;

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Select()
    {
        _walls?.Select();
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing)
        {

            if (_walls != null)
            {
                Marshal.ReleaseComObject(_walls);
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