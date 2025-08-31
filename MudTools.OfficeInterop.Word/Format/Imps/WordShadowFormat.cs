//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word.Imps;

/// <summary>
/// 封装 Microsoft.Office.Core.ShadowFormat 的实现类。
/// </summary>
internal class WordShadowFormat : IWordShadowFormat
{
    private MsWord.ShadowFormat _shadowFormat;
    private bool _disposedValue;

    /// <summary>
    /// 构造函数，包装 COM 对象。
    /// </summary>
    /// <param name="shadowFormat">原始 COM ShadowFormat 对象。</param>
    internal WordShadowFormat(MsWord.ShadowFormat shadowFormat)
    {
        _shadowFormat = shadowFormat ?? throw new ArgumentNullException(nameof(shadowFormat));
        _disposedValue = false;
    }

    #region 属性实现

    /// <inheritdoc/>
    public IWordApplication? Application => _shadowFormat != null ? new WordApplication(_shadowFormat.Application) : null;

    /// <inheritdoc/>
    public object Parent => _shadowFormat?.Parent;

    /// <inheritdoc/>
    public IWordColorFormat ForeColor
    {
        get => _shadowFormat?.ForeColor != null ? new WordColorFormat(_shadowFormat.ForeColor) : null;
    }

    /// <inheritdoc/>
    public float Transparency
    {
        get => _shadowFormat?.Transparency ?? 0f;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Transparency = value;
        }
    }

    /// <inheritdoc/>
    public bool Visible
    {
        get => _shadowFormat?.Visible == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Visible = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public MsoShadowType Type
    {
        get => _shadowFormat?.Type != null ? (MsoShadowType)(int)_shadowFormat?.Type : MsoShadowType.msoShadowMixed;
        set
        {
            if (_shadowFormat != null) _shadowFormat.Type = (MsCore.MsoShadowType)(int)value;
        }
    }

    /// <inheritdoc/>
    public float Blur
    {
        get => _shadowFormat?.Blur ?? 0f;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Blur = value;
        }
    }

    /// <inheritdoc/>
    public float Size
    {
        get => _shadowFormat?.Size ?? 100f;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.Size = value;
        }
    }

    /// <inheritdoc/>
    public bool RotateWithShape
    {
        get => _shadowFormat?.RotateWithShape == MsCore.MsoTriState.msoTrue;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.RotateWithShape = value ? MsCore.MsoTriState.msoTrue : MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public MsoShadowStyle Style
    {
        get => _shadowFormat?.Style != null ? (MsoShadowStyle)(int)_shadowFormat?.Style : MsoShadowStyle.msoShadowStyleMixed;
        set
        {
            if (_shadowFormat != null) _shadowFormat.Style = (MsCore.MsoShadowStyle)(int)value;
        }
    }

    /// <inheritdoc/>
    public float OffsetX
    {
        get => _shadowFormat?.OffsetX ?? 0f;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.OffsetX = value;
        }
    }

    /// <inheritdoc/>
    public float OffsetY
    {
        get => _shadowFormat?.OffsetY ?? 0f;
        set
        {
            if (_shadowFormat != null)
                _shadowFormat.OffsetY = value;
        }
    }

    #endregion

    #region 方法实现   

    /// <inheritdoc/>
    public void SetOffset(float offsetX, float offsetY)
    {
        if (_shadowFormat != null)
        {
            _shadowFormat.OffsetX = offsetX;
            _shadowFormat.OffsetY = offsetY;
        }
    }

    /// <inheritdoc/>
    public void Clear()
    {
        if (_shadowFormat != null)
        {
            _shadowFormat.Visible = MsCore.MsoTriState.msoFalse;
        }
    }

    /// <inheritdoc/>
    public void CopyTo(IWordShadowFormat targetShadow)
    {
        if (_shadowFormat == null || targetShadow == null)
            return;

        try
        {
            targetShadow.Visible = this.Visible;
            targetShadow.Type = this.Type;
            targetShadow.Blur = this.Blur;
            targetShadow.Size = this.Size;
            targetShadow.Transparency = this.Transparency;
            targetShadow.Style = this.Style;
            targetShadow.OffsetX = this.OffsetX;
            targetShadow.OffsetY = this.OffsetY;
            targetShadow.RotateWithShape = this.RotateWithShape;

            // 复制颜色
            if (this.ForeColor != null && targetShadow.ForeColor != null)
            {
                targetShadow.ForeColor.RGB = this.ForeColor.RGB;
            }
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException("无法复制阴影格式。", ex);
        }
    }

    /// <inheritdoc/>
    public void Reset()
    {
        if (_shadowFormat != null)
        {
            _shadowFormat.Visible = 0;
            _shadowFormat.Type = MsCore.MsoShadowType.msoShadow1;
            _shadowFormat.Blur = 0f;
            _shadowFormat.Size = 100f;
            _shadowFormat.Transparency = 0f;
            _shadowFormat.Style = MsCore.MsoShadowStyle.msoShadowStyleOuterShadow;
            _shadowFormat.OffsetX = 0f;
            _shadowFormat.OffsetY = 0f;
            _shadowFormat.RotateWithShape = MsCore.MsoTriState.msoTrue;
        }
    }

    /// <inheritdoc/>
    public void ApplyOuterShadow(float offsetX, float offsetY, float blur, int color, float transparency)
    {
        if (_shadowFormat != null)
        {
            _shadowFormat.Style = MsCore.MsoShadowStyle.msoShadowStyleOuterShadow;
            _shadowFormat.OffsetX = offsetX;
            _shadowFormat.OffsetY = offsetY;
            _shadowFormat.Blur = blur;
            _shadowFormat.Transparency = transparency;
            _shadowFormat.Visible = MsCore.MsoTriState.msoTrue;

            if (_shadowFormat.ForeColor != null)
                _shadowFormat.ForeColor.RGB = color;
        }
    }

    /// <inheritdoc/>
    public void ApplyInnerShadow(float offsetX, float offsetY, float blur, int color, float transparency)
    {
        if (_shadowFormat != null)
        {
            _shadowFormat.Style = MsCore.MsoShadowStyle.msoShadowStyleInnerShadow;
            _shadowFormat.OffsetX = offsetX;
            _shadowFormat.OffsetY = offsetY;
            _shadowFormat.Blur = blur;
            _shadowFormat.Transparency = transparency;
            _shadowFormat.Visible = MsCore.MsoTriState.msoTrue;

            if (_shadowFormat.ForeColor != null)
                _shadowFormat.ForeColor.RGB = color;
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

        if (disposing)
        {
            // 释放前景色对象
            if (_shadowFormat?.ForeColor != null)
            {
                Marshal.ReleaseComObject(_shadowFormat.ForeColor);
            }
            // 释放阴影格式对象本身
            if (_shadowFormat != null)
            {
                Marshal.ReleaseComObject(_shadowFormat);
                _shadowFormat = null;
            }
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