//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;
/// <summary>
/// Word.Interior 的封装实现类。
/// </summary>
internal class WordInterior : IWordInterior
{
    private MsWord.Interior _interior;
    private bool _disposedValue;

    internal WordInterior(MsWord.Interior interior)
    {
        _interior = interior ?? throw new ArgumentNullException(nameof(interior));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _interior != null ? new WordApplication(_interior.Application as MsWord.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _interior?.Parent;

    /// <inheritdoc/>
    public object Color
    {
        get => _interior?.Color;
        set
        {
            if (_interior != null)
                _interior.Color = value;
        }
    }

    /// <inheritdoc/>
    public XlColorIndex ColorIndex
    {
        get => _interior?.ColorIndex != null ? (XlColorIndex)(int)_interior?.ColorIndex : XlColorIndex.xlColorIndexNone;
        set
        {
            if (_interior != null) _interior.ColorIndex = (MsWord.XlColorIndex)(int)value;
        }
    }

    /// <inheritdoc/>
    public XlPattern Pattern
    {
        get => _interior?.Pattern != null ? (XlPattern)(int)_interior?.Pattern : XlPattern.xlPatternNone;
        set
        {
            if (_interior != null) _interior.Pattern = (MsWord.XlPattern)(int)value;
        }
    }

    /// <inheritdoc/>
    public object PatternColor
    {
        get => _interior?.PatternColor;
        set
        {
            if (_interior != null)
                _interior.PatternColor = value;
        }
    }

    /// <inheritdoc/>
    public XlColorIndex PatternColorIndex
    {
        get => _interior?.PatternColorIndex != null ? (XlColorIndex)(int)_interior?.PatternColorIndex : XlColorIndex.xlColorIndexNone;
        set
        {
            if (_interior != null) _interior.PatternColorIndex = (MsWord.XlColorIndex)(int)value;
        }
    }
    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _interior != null)
        {
            Marshal.ReleaseComObject(_interior);
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