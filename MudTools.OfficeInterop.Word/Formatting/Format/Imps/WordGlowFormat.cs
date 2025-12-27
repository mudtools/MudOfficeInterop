//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

// <summary>
/// Word.GlowFormat 的封装实现类。
/// </summary>
internal class WordGlowFormat : IWordGlowFormat
{
    private MsWord.GlowFormat _glowFormat;
    private bool _disposedValue;

    internal WordGlowFormat(MsWord.GlowFormat glowFormat)
    {
        _glowFormat = glowFormat ?? throw new ArgumentNullException(nameof(glowFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _glowFormat != null ? new WordApplication(_glowFormat.Application) : null;

    /// <inheritdoc/>
    public object? Parent => _glowFormat?.Parent;

    /// <inheritdoc/>
    public IWordColorFormat? Color
    {
        get
        {
            if (_glowFormat?.Color != null)
                return new WordColorFormat(_glowFormat.Color);
            return null;
        }
    }

    /// <inheritdoc/>
    public float Radius
    {
        get => _glowFormat?.Radius ?? 0f;
        set
        {
            if (_glowFormat != null)
                _glowFormat.Radius = value;
        }
    }

    /// <inheritdoc/>
    public float Transparency
    {
        get => _glowFormat?.Transparency ?? 0f;
        set
        {
            if (_glowFormat != null)
                _glowFormat.Transparency = value;
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _glowFormat != null)
        {
            Marshal.ReleaseComObject(_glowFormat);
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