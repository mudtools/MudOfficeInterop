//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// Word.ReflectionFormat 的封装实现类。
/// </summary>
internal class WordReflectionFormat : IWordReflectionFormat
{
    private MsWord.ReflectionFormat _reflectionFormat;
    private bool _disposedValue;

    internal WordReflectionFormat(MsWord.ReflectionFormat reflectionFormat)
    {
        _reflectionFormat = reflectionFormat ?? throw new ArgumentNullException(nameof(reflectionFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication Application => _reflectionFormat != null ? new WordApplication(_reflectionFormat.Application) : null;

    /// <inheritdoc/>
    public object Parent => _reflectionFormat?.Parent;

    /// <inheritdoc/>
    public MsoReflectionType Type
    {
        get => _reflectionFormat?.Type != null ? (MsoReflectionType)(int)_reflectionFormat?.Type : MsoReflectionType.msoReflectionTypeNone;
        set
        {
            if (_reflectionFormat != null) _reflectionFormat.Type = (MsCore.MsoReflectionType)(int)value;
        }
    }

    /// <inheritdoc/>
    public float Transparency
    {
        get => _reflectionFormat?.Transparency ?? 0f;
        set
        {
            if (_reflectionFormat != null)
                _reflectionFormat.Transparency = value;
        }
    }

    /// <inheritdoc/>
    public float Size
    {
        get => _reflectionFormat?.Size ?? 0f;
        set
        {
            if (_reflectionFormat != null)
                _reflectionFormat.Size = value;
        }
    }

    /// <inheritdoc/>
    public float Offset
    {
        get => _reflectionFormat?.Offset ?? 0f;
        set
        {
            if (_reflectionFormat != null)
                _reflectionFormat.Offset = value;
        }
    }

    /// <inheritdoc/>
    public float Blur
    {
        get => _reflectionFormat?.Blur ?? 0f;
        set
        {
            if (_reflectionFormat != null)
                _reflectionFormat.Blur = value;
        }
    }

    #endregion

    #region IDisposable 实现

    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _reflectionFormat != null)
        {
            Marshal.ReleaseComObject(_reflectionFormat);
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