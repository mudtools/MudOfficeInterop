//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Interop.Word.Shading 的实现类。
/// </summary>
internal class WordShading : IWordShading
{
    private MsWord.Shading _shading;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="shading">原始 COM Shading 对象。</param>
    internal WordShading(MsWord.Shading shading)
    {
        _shading = shading ?? throw new ArgumentNullException(nameof(shading));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _shading != null ? new WordApplication(_shading.Application) : null;

    /// <inheritdoc/>
    public object Parent => _shading?.Parent;

    /// <inheritdoc/>
    public WdColor BackgroundPatternColor
    {
        get => (WdColor)(int)_shading?.BackgroundPatternColor;
        set
        {
            if (_shading != null)
                _shading.BackgroundPatternColor = (MsWord.WdColor)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdColor ForegroundPatternColor
    {
        get => (WdColor)(int)_shading?.ForegroundPatternColor;
        set
        {
            if (_shading != null)
                _shading.ForegroundPatternColor = (MsWord.WdColor)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdTextureIndex Texture
    {
        get => (WdTextureIndex)(int)_shading?.Texture;
        set
        {
            if (_shading != null)
                _shading.Texture = (MsWord.WdTextureIndex)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdColorIndex BackgroundPatternColorIndex
    {
        get => (WdColorIndex)(int)_shading?.BackgroundPatternColorIndex;
        set
        {
            if (_shading != null)
                _shading.BackgroundPatternColorIndex = (MsWord.WdColorIndex)(int)value;
        }
    }

    /// <inheritdoc/>
    public WdColorIndex ForegroundPatternColorIndex
    {
        get => (WdColorIndex)(int)_shading?.ForegroundPatternColorIndex;
        set
        {
            if (_shading != null)
                _shading.ForegroundPatternColorIndex = (MsWord.WdColorIndex)(int)value;
        }
    }

    #endregion

    #region 方法实现

    /// <inheritdoc/>
    public void Clear()
    {
        if (_shading != null)
        {
            _shading.Texture = MsWord.WdTextureIndex.wdTextureNone;
            _shading.BackgroundPatternColor = MsWord.WdColor.wdColorWhite;
            _shading.ForegroundPatternColor = MsWord.WdColor.wdColorWhite;
        }
    }

    /// <inheritdoc/>
    public void ApplySolidColor(WdColor color)
    {
        if (_shading == null)
            throw new ObjectDisposedException(nameof(WordShading));

        try
        {
            _shading.Texture = MsWord.WdTextureIndex.wdTextureNone;
            _shading.BackgroundPatternColor = (MsWord.WdColor)(int)color;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法应用纯色底纹。", ex);
        }
    }

    /// <inheritdoc/>
    public void ApplyTexture(WdTextureIndex texture)
    {
        if (_shading == null)
            throw new ObjectDisposedException(nameof(WordShading));

        _shading.Texture = (MsWord.WdTextureIndex)(int)texture;
    }

    /// <inheritdoc/>
    public void CopyTo(IWordShading targetShading)
    {
        if (_shading == null || targetShading == null)
            return;

        try
        {
            targetShading.Texture = this.Texture;
            targetShading.BackgroundPatternColor = this.BackgroundPatternColor;
            targetShading.ForegroundPatternColor = this.ForegroundPatternColor;
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制底纹设置。", ex);
        }
    }

    #endregion

    #region IDisposable 实现

    /// <summary>
    /// 释放 COM 对象资源。
    /// </summary>
    /// <param name="disposing">是否由用户主动调用 Dispose。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposedValue) return;

        if (disposing && _shading != null)
        {
            Marshal.ReleaseComObject(_shading);
            _shading = null;
        }

        _disposedValue = true;
    }

    /// <inheritdoc/>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    #endregion
}